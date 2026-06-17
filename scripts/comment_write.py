#!/usr/bin/env python3
"""
comment_write.py — Cross-platform Word comment write operations (no word-mcp-live, no Word).

WHY THIS EXISTS
---------------
The plugin's designed write path for comments (reply / resolve / delete) goes
through the `word-mcp-live` MCP server, which: (a) must be separately installed,
(b) on macOS drives Word via JXA with reduced capability, (c) on Windows needs
Word + COM. None of that is guaranteed to be present.

This script performs the same operations by editing the OOXML parts directly
(`word/comments.xml`, `word/commentsExtended.xml`, `word/document.xml`). It is
pure python-docx/lxml + zipfile, so it behaves identically on macOS, Windows,
and Linux and needs neither word-mcp-live nor Microsoft Word.

Pairs with `extract_comments.py` (the read path).

Operations:
    reply      --id N --text "..." [--author NAME] [--initials XX]
    resolve    --id N
    unresolve  --id N
    delete     --id N            (also removes the thread's replies)
    delete     --author NAME     (delete all comments by an author)
    delete-all                   (strip every comment)
    list                         (id/author/resolved overview; convenience)

Usage:
    python3 comment_write.py FILE.docx OP [args]
    Options: --output PATH (default: in place) ; --no-backup

Output: JSON {success, operation, ...} to stdout. Exit 0 on success, 1 on error.
"""
import sys
import os
import json
import time
import uuid
import shutil
import zipfile
import argparse
import tempfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS = {'w': W, 'w14': W14, 'w15': W15}

CT_PATH = '[Content_Types].xml'
RELS_PATH = 'word/_rels/document.xml.rels'
COMMENTS_PATH = 'word/comments.xml'
EXTENDED_PATH = 'word/commentsExtended.xml'
DOC_PATH = 'word/document.xml'

CT_EXTENDED = ('<Override PartName="/word/commentsExtended.xml" '
               'ContentType="application/vnd.openxmlformats-officedocument.'
               'wordprocessingml.commentsExtended+xml"/>')
REL_EXTENDED_TYPE = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'


def _q(ns, tag):
    return f"{{{ns}}}{tag}"


def _new_paraid():
    """8 uppercase hex digits, as Word uses for w14:paraId."""
    return uuid.uuid4().hex[:8].upper()


class Docx:
    """Minimal in-memory docx editor over the zipped OOXML parts."""

    def __init__(self, path):
        self.path = path
        self.tmp = tempfile.mkdtemp(prefix='cmtwrite_')
        with zipfile.ZipFile(path) as z:
            z.extractall(self.tmp)
            self._names = z.namelist()
        self._trees = {}

    def cleanup(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _abs(self, rel):
        return os.path.join(self.tmp, rel.replace('/', os.sep))

    def has(self, rel):
        return os.path.exists(self._abs(rel))

    def tree(self, rel):
        if rel not in self._trees:
            self._trees[rel] = etree.parse(self._abs(rel))
        return self._trees[rel]

    def root(self, rel):
        return self.tree(rel).getroot()

    def read_text(self, rel):
        with open(self._abs(rel), encoding='utf-8') as f:
            return f.read()

    def write_text(self, rel, text):
        with open(self._abs(rel), 'w', encoding='utf-8') as f:
            f.write(text)

    def save_tree(self, rel):
        self.tree(rel).write(self._abs(rel), xml_declaration=True,
                             encoding='UTF-8', standalone=True)

    def repack(self, out_path):
        # Write [Content_Types].xml first (good practice), then the rest.
        files = []
        for r, _, fs in os.walk(self.tmp):
            for f in fs:
                full = os.path.join(r, f)
                arc = os.path.relpath(full, self.tmp).replace(os.sep, '/')
                files.append((arc, full))
        files.sort(key=lambda t: (t[0] != CT_PATH, t[0]))
        tmp_out = out_path + '.tmp'
        with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as z:
            for arc, full in files:
                z.write(full, arc)
        os.replace(tmp_out, out_path)


# ---------- shared lookups ----------

def _comments_root(dx):
    if not dx.has(COMMENTS_PATH):
        return None
    return dx.root(COMMENTS_PATH)


def _comment_by_id(croot, cid):
    for c in croot.findall(_q(W, 'comment')):
        if c.get(_q(W, 'id')) == str(cid):
            return c
    return None


def _max_comment_id(croot):
    mx = -1
    for c in croot.findall(_q(W, 'comment')):
        try:
            mx = max(mx, int(c.get(_q(W, 'id'))))
        except (TypeError, ValueError):
            pass
    return mx


def _ensure_paraid(comment_el):
    """Return the w14:paraId of a comment's first paragraph, creating one if absent."""
    p = comment_el.find(_q(W, 'p'))
    if p is None:
        p = etree.SubElement(comment_el, _q(W, 'p'))
    pid = p.get(_q(W14, 'paraId'))
    if not pid:
        pid = _new_paraid()
        p.set(_q(W14, 'paraId'), pid)
    return pid


def _extended_root(dx, create=False):
    """Return commentsExtended root, optionally creating the part + registration."""
    if dx.has(EXTENDED_PATH):
        return dx.root(EXTENDED_PATH)
    if not create:
        return None
    root = etree.Element(_q(W15, 'commentsEx'), nsmap={'w15': W15})
    dx._trees[EXTENDED_PATH] = etree.ElementTree(root)
    # register content type
    ct = dx.read_text(CT_PATH)
    if 'commentsExtended' not in ct:
        dx.write_text(CT_PATH, ct.replace('</Types>', CT_EXTENDED + '</Types>'))
    # register relationship
    rels = dx.read_text(RELS_PATH)
    if 'commentsExtended' not in rels:
        rid = 'rIdCmtEx'
        rel = (f'<Relationship Id="{rid}" Type="{REL_EXTENDED_TYPE}" '
               f'Target="commentsExtended.xml"/>')
        dx.write_text(RELS_PATH, rels.replace('</Relationships>', rel + '</Relationships>'))
    return root


def _ext_entry(eroot, paraid, create=False):
    for ex in eroot.findall(_q(W15, 'commentEx')):
        if ex.get(_q(W15, 'paraId')) == paraid:
            return ex
    if not create:
        return None
    ex = etree.SubElement(eroot, _q(W15, 'commentEx'))
    ex.set(_q(W15, 'paraId'), paraid)
    ex.set(_q(W15, 'done'), '0')
    return ex


def _paraid_to_cid(croot):
    """Map each comment's first-paragraph paraId -> comment id."""
    out = {}
    for c in croot.findall(_q(W, 'comment')):
        p = c.find(_q(W, 'p'))
        if p is not None:
            pid = p.get(_q(W14, 'paraId'))
            if pid:
                out[pid] = c.get(_q(W, 'id'))
    return out


# ---------- operations ----------

def op_reply(dx, target_id, text, author, initials):
    croot = _comments_root(dx)
    if croot is None:
        raise ValueError('Document has no comments to reply to.')
    parent = _comment_by_id(croot, target_id)
    if parent is None:
        raise ValueError(f'Comment id {target_id} not found.')
    parent_pid = _ensure_paraid(parent)

    new_id = str(_max_comment_id(croot) + 1)
    reply_pid = _new_paraid()
    # new <w:comment> in comments.xml
    rc = etree.SubElement(croot, _q(W, 'comment'))
    rc.set(_q(W, 'id'), new_id)
    rc.set(_q(W, 'author'), author)
    rc.set(_q(W, 'initials'), initials or '')
    rc.set(_q(W, 'date'), time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()))
    rp = etree.SubElement(rc, _q(W, 'p'))
    rp.set(_q(W14, 'paraId'), reply_pid)
    rr = etree.SubElement(rp, _q(W, 'r'))
    rt = etree.SubElement(rr, _q(W, 't'))
    rt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    rt.text = text

    # thread link in commentsExtended.xml
    eroot = _extended_root(dx, create=True)
    _ext_entry(eroot, parent_pid, create=True)  # ensure parent exists
    rex = _ext_entry(eroot, reply_pid, create=True)
    rex.set(_q(W15, 'paraIdParent'), parent_pid)

    # commentReference run in document.xml, right after the parent's reference,
    # so Word renders the reply in the same thread/anchor.
    droot = dx.root(DOC_PATH)
    ref = None
    for r in droot.iter(_q(W, 'commentReference')):
        if r.get(_q(W, 'id')) == str(target_id):
            ref = r
            break
    if ref is not None:
        run = ref.getparent()                 # the <w:r> wrapping the reference
        new_run = etree.SubElement(run.getparent(), _q(W, 'r'))
        run.addnext(new_run)
        nref = etree.SubElement(new_run, _q(W, 'commentReference'))
        nref.set(_q(W, 'id'), new_id)

    dx.save_tree(COMMENTS_PATH)
    dx.save_tree(EXTENDED_PATH)
    dx.save_tree(DOC_PATH)
    return {'reply_id': new_id, 'parent_id': str(target_id)}


def _set_done(dx, target_id, done):
    croot = _comments_root(dx)
    if croot is None:
        raise ValueError('Document has no comments.')
    c = _comment_by_id(croot, target_id)
    if c is None:
        raise ValueError(f'Comment id {target_id} not found.')
    pid = _ensure_paraid(c)
    eroot = _extended_root(dx, create=True)
    ex = _ext_entry(eroot, pid, create=True)
    ex.set(_q(W15, 'done'), '1' if done else '0')
    dx.save_tree(COMMENTS_PATH)   # paraId may have been added
    dx.save_tree(EXTENDED_PATH)
    return {'comment_id': str(target_id), 'resolved': bool(done)}


def op_resolve(dx, target_id):
    return _set_done(dx, target_id, True)


def op_unresolve(dx, target_id):
    return _set_done(dx, target_id, False)


def _delete_ids(dx, ids):
    """Remove the given comment ids (and their replies) from all parts."""
    croot = _comments_root(dx)
    if croot is None:
        return {'deleted': []}
    ids = set(str(i) for i in ids)

    # Expand to include replies: any comment whose paraIdParent chains to a target.
    eroot = _extended_root(dx, create=False)
    if eroot is not None:
        p2c = _paraid_to_cid(croot)
        cid2pid = {v: k for k, v in p2c.items()}
        # parent paraId -> child comment ids
        changed = True
        while changed:
            changed = False
            for ex in eroot.findall(_q(W15, 'commentEx')):
                parent_pid = ex.get(_q(W15, 'paraIdParent'))
                child_pid = ex.get(_q(W15, 'paraId'))
                if parent_pid and parent_pid in {cid2pid.get(i) for i in ids}:
                    child_cid = p2c.get(child_pid)
                    if child_cid and child_cid not in ids:
                        ids.add(child_cid)
                        changed = True

    # collect paraIds of doomed comments before removing
    doomed_pids = set()
    for c in list(croot.findall(_q(W, 'comment'))):
        if c.get(_q(W, 'id')) in ids:
            p = c.find(_q(W, 'p'))
            if p is not None and p.get(_q(W14, 'paraId')):
                doomed_pids.add(p.get(_q(W14, 'paraId')))
            croot.remove(c)

    # remove range markers + references in document.xml
    droot = dx.root(DOC_PATH)
    for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
        for el in list(droot.iter(_q(W, tag))):
            if el.get(_q(W, 'id')) in ids:
                parent = el.getparent()
                # if the <w:r> wrapper now only held this reference, drop it too
                if (tag == 'commentReference' and parent.tag == _q(W, 'r')
                        and len(parent) == 1):
                    gp = parent.getparent()
                    if gp is not None:
                        gp.remove(parent)
                        continue
                parent.remove(el)

    # remove commentsExtended entries
    if eroot is not None:
        for ex in list(eroot.findall(_q(W15, 'commentEx'))):
            if ex.get(_q(W15, 'paraId')) in doomed_pids:
                eroot.remove(ex)
        dx.save_tree(EXTENDED_PATH)

    dx.save_tree(COMMENTS_PATH)
    dx.save_tree(DOC_PATH)
    return {'deleted': sorted(ids, key=lambda x: int(x) if x.isdigit() else 0)}


def op_delete(dx, target_id=None, author=None):
    croot = _comments_root(dx)
    if croot is None:
        return {'deleted': []}
    if author is not None:
        a = author.lower()
        ids = [c.get(_q(W, 'id')) for c in croot.findall(_q(W, 'comment'))
               if c.get(_q(W, 'author'), '').lower() == a]
    else:
        ids = [str(target_id)]
    return _delete_ids(dx, ids)


def op_delete_all(dx):
    croot = _comments_root(dx)
    if croot is None:
        return {'deleted': []}
    ids = [c.get(_q(W, 'id')) for c in croot.findall(_q(W, 'comment'))]
    return _delete_ids(dx, ids)


def op_list(dx):
    croot = _comments_root(dx)
    if croot is None:
        return {'comments': []}
    eroot = _extended_root(dx, create=False)
    done = {}
    if eroot is not None:
        for ex in eroot.findall(_q(W15, 'commentEx')):
            done[ex.get(_q(W15, 'paraId'))] = ex.get(_q(W15, 'done')) in ('1', 'true')
    out = []
    for c in croot.findall(_q(W, 'comment')):
        p = c.find(_q(W, 'p'))
        pid = p.get(_q(W14, 'paraId')) if p is not None else None
        texts = c.xpath('.//w:t', namespaces=NS)
        out.append({
            'comment_id': c.get(_q(W, 'id')),
            'author': c.get(_q(W, 'author'), 'Unknown'),
            'resolved': done.get(pid, False),
            'text': ''.join(t.text or '' for t in texts).strip(),
        })
    return {'comments': out}


def main(argv=None):
    ap = argparse.ArgumentParser(description='Write Word comment operations via pure OOXML editing.')
    ap.add_argument('file')
    sub = ap.add_subparsers(dest='op', required=True)

    pr = sub.add_parser('reply')
    pr.add_argument('--id', type=int, required=True)
    pr.add_argument('--text', required=True)
    pr.add_argument('--author', default=os.environ.get('MCP_AUTHOR', 'Author'))
    pr.add_argument('--initials', default='')

    for name in ('resolve', 'unresolve'):
        sp = sub.add_parser(name)
        sp.add_argument('--id', type=int, required=True)

    pd = sub.add_parser('delete')
    pd.add_argument('--id', type=int)
    pd.add_argument('--author')

    sub.add_parser('delete-all')
    sub.add_parser('list')

    for sp in sub.choices.values():
        sp.add_argument('--output')
        sp.add_argument('--no-backup', action='store_true')

    args = ap.parse_args(argv)

    path = args.file if args.file.lower().endswith('.docx') else args.file + '.docx'
    if not os.path.exists(path):
        print(json.dumps({'success': False, 'error': f'File not found: {path}'}, ensure_ascii=False, indent=2))
        return 1

    if args.op == 'delete' and args.id is None and not args.author:
        print(json.dumps({'success': False, 'error': 'delete requires --id or --author'}, ensure_ascii=False, indent=2))
        return 1

    dx = Docx(path)
    try:
        if args.op == 'reply':
            result = op_reply(dx, args.id, args.text, args.author, args.initials)
        elif args.op == 'resolve':
            result = op_resolve(dx, args.id)
        elif args.op == 'unresolve':
            result = op_unresolve(dx, args.id)
        elif args.op == 'delete':
            result = op_delete(dx, target_id=args.id, author=args.author)
        elif args.op == 'delete-all':
            result = op_delete_all(dx)
        elif args.op == 'list':
            print(json.dumps({'success': True, 'operation': 'list', **op_list(dx)},
                             ensure_ascii=False, indent=2))
            return 0
        else:
            raise ValueError(f'Unknown op {args.op}')

        out_path = args.output or path
        if out_path == path and not args.no_backup:
            bak = f"{os.path.splitext(path)[0]}_backup_{time.strftime('%Y%m%d_%H%M%S')}.docx"
            shutil.copy2(path, bak)
            result['backup'] = bak
        dx.repack(out_path)
        result['output'] = out_path
        print(json.dumps({'success': True, 'operation': args.op, **result},
                         ensure_ascii=False, indent=2))
        return 0
    except Exception as e:
        print(json.dumps({'success': False, 'operation': args.op, 'error': str(e)},
                         ensure_ascii=False, indent=2))
        return 1
    finally:
        dx.cleanup()


if __name__ == '__main__':
    sys.exit(main())

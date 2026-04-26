#!/bin/bash
# Word-Agent Setup Script
# Checks dependencies, installs Python packages, and registers the Claude Code plugin.

set -e

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

pass() { echo -e "${GREEN}✓${NC} $1"; }
fail() { echo -e "${RED}✗${NC} $1"; }
warn() { echo -e "${YELLOW}!${NC} $1"; }
info() { echo -e "  $1"; }

echo "==============================="
echo "  Word-Agent Setup"
echo "==============================="
echo ""

ERRORS=0

# 1. Check Python
echo "Checking dependencies..."
if command -v python3 &>/dev/null; then
    PY_VERSION=$(python3 --version 2>&1)
    pass "Python3: $PY_VERSION"
else
    fail "Python3 not found. Please install Python 3.8+"
    ERRORS=$((ERRORS + 1))
fi

# 2. Check pip
if command -v pip3 &>/dev/null; then
    pass "pip3 available"
else
    fail "pip3 not found. Please install pip"
    ERRORS=$((ERRORS + 1))
fi

# 3. Check Claude Code
if command -v claude &>/dev/null; then
    pass "Claude Code CLI available"
else
    fail "Claude Code CLI not found. Install from: https://docs.anthropic.com/en/docs/claude-code"
    ERRORS=$((ERRORS + 1))
fi

if [ $ERRORS -gt 0 ]; then
    echo ""
    fail "Found $ERRORS missing dependencies. Please install them and re-run."
    exit 1
fi

# 4. Install Python dependencies
echo ""
echo "Installing Python dependencies..."
if pip3 install -r "$SCRIPT_DIR/requirements.txt" --quiet 2>&1; then
    pass "Python packages installed"
else
    fail "Failed to install Python packages"
    exit 1
fi

# 5. Verify key packages
echo ""
echo "Verifying packages..."
for pkg in docx2python deepdiff redlines; do
    if python3 -c "import $pkg" 2>/dev/null; then
        pass "$pkg"
    else
        fail "$pkg import failed"
        ERRORS=$((ERRORS + 1))
    fi
done

# 6. Register marketplace and install plugin
echo ""
echo "Installing Claude Code plugin..."

# Uninstall if already installed (to refresh cache)
claude plugin uninstall word-agent@word-agent 2>/dev/null && info "Removed old version"

# Add marketplace (idempotent)
if claude plugin marketplace add "$PROJECT_DIR" 2>&1 | grep -qiE "success|already"; then
    pass "Marketplace registered"
else
    # marketplace add might show various messages, try anyway
    info "Marketplace registration attempted"
fi

# Install plugin
if claude plugin install word-agent@word-agent 2>&1 | grep -qi "success"; then
    pass "Plugin installed"
else
    fail "Plugin installation failed"
    info "Try manually: claude plugin marketplace add \"$PROJECT_DIR\""
    info "             claude plugin install word-agent@word-agent"
    exit 1
fi

# 7. Check MCP server
echo ""
echo "Checking MCP servers..."
if claude mcp list 2>/dev/null | grep -qi "word-document-server"; then
    pass "word-document-server MCP configured"
else
    warn "word-document-server MCP not found"
    info "word-agent requires this MCP server to function."
    info "Please configure it in your Claude Code MCP settings."
fi

if claude mcp list 2>/dev/null | grep -qi "zotero"; then
    pass "Zotero MCP configured (optional)"
else
    info "Zotero MCP not configured (optional, for reference management)"
fi

# Done
echo ""
echo "==============================="
pass "Word-Agent setup complete!"
echo "==============================="
echo ""
echo "Usage: restart Claude Code, then try:"
echo "  /word-agent:word-read"
echo "  分析一下这个文档：论文.docx"
echo ""

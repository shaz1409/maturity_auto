#!/bin/bash
# Helper script to set SharePoint environment variables

echo "🔧 SharePoint Environment Variables Setup"
echo "=========================================="
echo ""

# Check if .zshrc exists
ZSHRC_FILE="$HOME/.zshrc"

# Function to add to .zshrc
add_to_zshrc() {
    echo "" >> "$ZSHRC_FILE"
    echo "# SharePoint App Authentication (added $(date +%Y-%m-%d))" >> "$ZSHRC_FILE"
    echo "export SHAREPOINT_UPLOAD=true" >> "$ZSHRC_FILE"
    echo "export SHAREPOINT_AUTH_METHOD=app" >> "$ZSHRC_FILE"
    echo "export SHAREPOINT_CLIENT_ID='$1'" >> "$ZSHRC_FILE"
    echo "export SHAREPOINT_CLIENT_SECRET='$2'" >> "$ZSHRC_FILE"
}

# Prompt for values
read -p "Enter your SharePoint Client ID: " CLIENT_ID
read -sp "Enter your SharePoint Client Secret: " CLIENT_SECRET
echo ""

# Set for current session
export SHAREPOINT_UPLOAD=true
export SHAREPOINT_AUTH_METHOD=app
export SHAREPOINT_CLIENT_ID="$CLIENT_ID"
export SHAREPOINT_CLIENT_SECRET="$CLIENT_SECRET"

echo ""
echo "✅ Environment variables set for current session"
echo ""

# Ask if they want to make it persistent
read -p "Add to ~/.zshrc for persistence? (y/n): " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    # Remove old entries if they exist
    if grep -q "SHAREPOINT_CLIENT_ID" "$ZSHRC_FILE" 2>/dev/null; then
        echo "⚠️  Found existing SharePoint variables in ~/.zshrc"
        read -p "Replace them? (y/n): " -n 1 -r
        echo ""
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            # Remove old SharePoint entries
            sed -i '' '/# SharePoint App Authentication/,/export SHAREPOINT_CLIENT_SECRET/d' "$ZSHRC_FILE" 2>/dev/null
        fi
    fi
    
    add_to_zshrc "$CLIENT_ID" "$CLIENT_SECRET"
    echo "✅ Added to ~/.zshrc"
    echo "   Run 'source ~/.zshrc' or open a new terminal to apply"
fi

echo ""
echo "📋 Current values:"
echo "   SHAREPOINT_UPLOAD: $SHAREPOINT_UPLOAD"
echo "   SHAREPOINT_AUTH_METHOD: $SHAREPOINT_AUTH_METHOD"
echo "   SHAREPOINT_CLIENT_ID: $SHAREPOINT_CLIENT_ID"
echo "   SHAREPOINT_CLIENT_SECRET: [hidden]"
echo ""
echo "🧪 Test with: python maturity_assessment.py"


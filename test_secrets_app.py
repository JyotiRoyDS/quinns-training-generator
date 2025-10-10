import streamlit as st
import os

st.title("🔍 Secrets Debugging Tool")

st.markdown("---")

# Test 1: Check if secrets exist
st.header("Test 1: Secrets Attribute")
if hasattr(st, 'secrets'):
    st.success("✅ Secrets attribute EXISTS")
else:
    st.error("❌ Secrets attribute MISSING")
    st.stop()

# Test 2: List all secret keys
st.header("Test 2: Available Secret Keys")
try:
    keys = list(st.secrets.keys())
    if keys:
        st.success(f"✅ Found {len(keys)} secret(s)")
        for key in keys:
            st.write(f"- `{key}`")
    else:
        st.warning("⚠️ No secrets found (empty)")
except Exception as e:
    st.error(f"❌ Error accessing secrets: {e}")

# Test 3: Check for OPENAI_API_KEY
st.header("Test 3: OPENAI_API_KEY Check")
try:
    if 'OPENAI_API_KEY' in st.secrets:
        st.success("✅ OPENAI_API_KEY found in secrets!")
        
        # Get the key
        api_key = st.secrets['OPENAI_API_KEY']
        
        # Show details
        st.write(f"**Key length:** {len(api_key)} characters")
        st.write(f"**Starts with:** `{api_key[:15]}...`")
        st.write(f"**Ends with:** `...{api_key[-15:]}`")
        
        # Check if valid format
        if api_key.startswith('sk-'):
            st.success("✅ Key format looks valid (starts with 'sk-')")
        else:
            st.error("❌ Key format may be invalid (should start with 'sk-')")
            
    else:
        st.error("❌ OPENAI_API_KEY NOT found in secrets")
        st.info("Available keys: " + str(list(st.secrets.keys())))
except Exception as e:
    st.error(f"❌ Error checking OPENAI_API_KEY: {e}")

# Test 4: Try the get_api_key function logic
st.header("Test 4: get_api_key() Logic Test")

def test_get_api_key():
    """Test version of get_api_key"""
    try:
        if hasattr(st, 'secrets') and 'OPENAI_API_KEY' in st.secrets:
            api_key = st.secrets['OPENAI_API_KEY']
            if api_key and api_key.strip():
                return api_key.strip()
    except Exception as e:
        return f"ERROR: {e}"
    
    api_key = os.getenv('OPENAI_API_KEY')
    if api_key and api_key.strip():
        return api_key.strip()
    
    return None

result = test_get_api_key()

if result:
    if result.startswith("ERROR:"):
        st.error(f"❌ Function returned error: {result}")
    else:
        st.success("✅ Function successfully retrieved key!")
        st.write(f"**Returned key length:** {len(result)} characters")
else:
    st.error("❌ Function returned None - key not found")

st.markdown("---")

# Instructions
st.header("📋 How to Fix")

st.markdown("""
**If you see errors above, your secrets format is wrong!**

### Correct Format:
```toml
OPENAI_API_KEY = "sk-proj-your-full-key-here"
```

### Common Mistakes:
- ❌ No spaces: `OPENAI_API_KEY="key"`
- ❌ Single quotes: `OPENAI_API_KEY = 'key'`
- ❌ Line breaks in the middle of the key
- ❌ Extra spaces at the end

### Steps to Fix:
1. Go to Settings → Secrets
2. Delete ALL text
3. Paste the correct format (with spaces around `=`)
4. Click "Save changes"
5. Wait 1-2 minutes
6. Refresh this page
""")

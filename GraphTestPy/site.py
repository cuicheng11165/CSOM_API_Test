import requests
import json
import os

# Configuration - use relative path from script location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "Config")
TOKEN_FILE = os.path.join(CONFIG_DIR, "GraphAuthorization.txt")
HOST_FILE = os.path.join(CONFIG_DIR, "HostName.txt")

def get_token():
    """Read the Graph API token from the config file"""
    with open(TOKEN_FILE, 'r') as f:
        return f.read().strip()

def get_hostname():
    """Read the SharePoint hostname from the config file"""
    with open(HOST_FILE, 'r') as f:
        return f.read().strip()

def make_graph_request(endpoint, token, method='GET', data=None):
    """Make a request to Microsoft Graph API"""
    headers = {
        'Authorization': token if token.startswith('Bearer ') else f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    
    url = f"https://graph.microsoft.com/v1.0{endpoint}"
    
    print(f"\n{method} {url}")
    print(f"Authorization header: {headers['Authorization'][:50]}...")
    
    if method == 'GET':
        response = requests.get(url, headers=headers)
    elif method == 'POST':
        response = requests.post(url, headers=headers, json=data)
    elif method == 'PATCH':
        response = requests.patch(url, headers=headers, json=data)
    
    # Print detailed error info for 401
    if response.status_code == 401:
        print(f"\n⚠️  401 UNAUTHORIZED ERROR")
        print(f"Response Headers: {dict(response.headers)}")
        try:
            error_detail = response.json()
            print(f"Error Details: {json.dumps(error_detail, indent=2)}")
        except:
            print(f"Response Text: {response.text}")
        print("\nPossible causes:")
        print("1. Token expired - regenerate token using ExportGraphToken")
        print("2. Token missing required scopes (Sites.Read.All, Sites.ReadWrite.All)")
        print("3. Token format incorrect - should be in GraphAuthorization.txt")
    
    return response

def test_list_sites(token):
    """Test: List all sites in the organization"""
    print("\n" + "="*60)
    print("TEST: List Sites")
    print("="*60)
    
    response = make_graph_request("/sites?$top=10", token)
    
    if response.status_code == 200:
        sites = response.json()
        print(f"\nFound {len(sites.get('value', []))} sites:")
        for site in sites.get('value', []):
            print(f"  - {site.get('displayName')} ({site.get('webUrl')})")
            print(f"    ID: {site.get('id')}")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_get_root_site(token):
    """Test: Get root site"""
    print("\n" + "="*60)
    print("TEST: Get Root Site")
    print("="*60)
    
    response = make_graph_request("/sites/root", token)
    
    if response.status_code == 200:
        site = response.json()
        print(f"\nRoot Site: {site.get('displayName')}")
        print(f"  URL: {site.get('webUrl')}")
        print(f"  ID: {site.get('id')}")
        print(f"  Description: {site.get('description', 'N/A')}")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_get_site_by_hostname(token, hostname):
    """Test: Get site by hostname and path"""
    print("\n" + "="*60)
    print("TEST: Get Site by Hostname")
    print("="*60)
    
    # Get root site
    endpoint = f"/sites/{hostname}:/"
    response = make_graph_request(endpoint, token)
    
    if response.status_code == 200:
        site = response.json()
        print(f"\nSite: {site.get('displayName')}")
        print(f"  URL: {site.get('webUrl')}")
        print(f"  ID: {site.get('id')}")
        return site.get('id')
    else:
        print(f"Error: {response.status_code}")
        print(response.text)
        return None

def test_list_site_lists(token, site_id):
    """Test: List all lists in a site"""
    print("\n" + "="*60)
    print("TEST: List Site Lists")
    print("="*60)
    
    if not site_id:
        print("Skipping - no site ID provided")
        return
    
    response = make_graph_request(f"/sites/{site_id}/lists", token)
    
    if response.status_code == 200:
        lists = response.json()
        print(f"\nFound {len(lists.get('value', []))} lists:")
        for lst in lists.get('value', []):
            print(f"  - {lst.get('displayName')} (ID: {lst.get('id')})")
            print(f"    Template: {lst.get('list', {}).get('template', 'N/A')}")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_list_site_drives(token, site_id):
    """Test: List all document libraries (drives) in a site"""
    print("\n" + "="*60)
    print("TEST: List Site Drives")
    print("="*60)
    
    if not site_id:
        print("Skipping - no site ID provided")
        return
    
    response = make_graph_request(f"/sites/{site_id}/drives", token)
    
    if response.status_code == 200:
        drives = response.json()
        print(f"\nFound {len(drives.get('value', []))} drives:")
        for drive in drives.get('value', []):
            print(f"  - {drive.get('name')} (ID: {drive.get('id')})")
            print(f"    Type: {drive.get('driveType')}")
            print(f"    URL: {drive.get('webUrl')}")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_search_sites(token, query):
    """Test: Search for sites"""
    print("\n" + "="*60)
    print(f"TEST: Search Sites (query: '{query}')")
    print("="*60)
    
    response = make_graph_request(f"/sites?$search={query}", token)
    
    if response.status_code == 200:
        sites = response.json()
        print(f"\nFound {len(sites.get('value', []))} sites:")
        for site in sites.get('value', []):
            print(f"  - {site.get('displayName')} ({site.get('webUrl')})")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_get_site_analytics(token, site_id):
    """Test: Get site analytics"""
    print("\n" + "="*60)
    print("TEST: Get Site Analytics")
    print("="*60)
    
    if not site_id:
        print("Skipping - no site ID provided")
        return
    
    response = make_graph_request(f"/sites/{site_id}/analytics", token)
    
    if response.status_code == 200:
        analytics = response.json()
        print("\nSite Analytics:")
        print(json.dumps(analytics, indent=2))
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def test_get_agent_files(token, site_id):
    """Test: Get .agent files from SiteAssets/Copilots folder"""
    print("\n" + "="*60)
    print("TEST: Get Agent Files from SiteAssets/Copilots")
    print("="*60)
    
    if not site_id:
        print("Skipping - no site ID provided")
        return
    
    # First, check if the folder exists by getting all items in SiteAssets/Copilots
    print("\nStep 1: Checking if SiteAssets/Copilots folder exists...")
    check_endpoint = f"/sites/{site_id}/drive/root:/SiteAssets/Copilots"
    check_response = make_graph_request(check_endpoint, token)
    
    if check_response.status_code == 404:
        print("❌ Folder 'SiteAssets/Copilots' does not exist")
        print("   Trying alternative path: /sites/{site_id}/drive/root/children (listing root)")
        
        # List root to see what's available
        root_endpoint = f"/sites/{site_id}/drive/root/children"
        root_response = make_graph_request(root_endpoint, token)
        if root_response.status_code == 200:
            items = root_response.json().get('value', [])
            print(f"\n   Available folders/files in root:")
            for item in items[:10]:  # Show first 10
                item_type = "📁" if item.get('folder') else "📄"
                print(f"   {item_type} {item.get('name')}")
        return
    elif check_response.status_code != 200:
        print(f"Error checking folder: {check_response.status_code}")
        print(check_response.text)
        return
    
    print("✓ Folder exists")
    
    # Now try to get children with filter - NOTE: $filter might not be supported on this endpoint
    print("\nStep 2: Getting all files from the folder...")
    endpoint = f"/sites/{site_id}/drive/root:/SiteAssets/Copilots:/children"
    response = make_graph_request(endpoint, token)
    
    if response.status_code == 200:
        result = response.json()
        all_files = result.get('value', [])
        
        # Filter .agent files client-side (since server-side filter may not work)
        agent_files = [f for f in all_files if f.get('name', '').endswith('.agent')]
        
        print(f"\nFound {len(all_files)} total files, {len(agent_files)} .agent files:")
        for file in agent_files:
            print(f"  - {file.get('name')}")
            print(f"    ID: {file.get('id')}")
            print(f"    Size: {file.get('size')} bytes")
            print(f"    Modified: {file.get('lastModifiedDateTime')}")
            print(f"    URL: {file.get('webUrl')}")
            print()
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

def main():
    """Run all tests"""
    print("="*60)
    print("Microsoft Graph Sites API Test")
    print("="*60)
    
    # Get token and hostname
    try:
        token = get_token()
        hostname = get_hostname()
        print(f"Token loaded: {token[:20]}...")
        print(f"Hostname: {hostname}")
    except FileNotFoundError as e:
        print(f"Error: Could not find config file - {e}")
        return
    
    # Run tests
    # test_get_root_site(token)
    # test_list_sites(token)
    
    # Get specific site by hostname
    # site_id = test_get_site_by_hostname(token, hostname)
    site_id = '7f4b0ba1-4ade-4a15-9952-65912d599bea'
    # Additional tests with site ID
    if site_id:
        # test_list_site_lists(token, site_id)
        # test_list_site_drives(token, site_id)
        # test_get_site_analytics(token, site_id)
        test_get_agent_files(token, site_id)
    
    # Search test
    # test_search_sites(token, "site")
    
    print("\n" + "="*60)
    print("All tests completed!")
    print("="*60)

if __name__ == "__main__":
    main()

import requests
import json
import os

def get_token():
    """Read token from Config/ClientSecretToken.txt"""
    token_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Config', 'ClientSecretToken.txt')
    with open(token_path, 'r') as f:
        return f.read().strip()

def list_assistants():
    """
    List all assistants from Azure AI Foundry project
    
    Prerequisites:
    1. Token must be generated with scope: https://ai.azure.com/.default
       Run: cd ExportTokenByClientSecret && dotnet run (from bin/Debug/net8.0)
    
    2. Service Principal must have permission:
       Microsoft.CognitiveServices/accounts/AIServices/agents/read
       
    Note: Currently returns PermissionDenied. To fix:
    - Grant the service principal appropriate role in Azure AI Foundry
    - Roles: Cognitive Services User, Cognitive Services Contributor, or custom role
    """
    base_url = "https://agentpulsedevtest.services.ai.azure.com/api/projects/AgentPulseDevTest"
    token = get_token()
    
    # Try common assistant endpoints with different API versions
    endpoints = [
        (f"{base_url}/assistants", {"api-version": "v1"}),
    ]
    
    print(f"Azure AI Foundry - List Assistants Test")
    print(f"Endpoint: {base_url}")
    print(f"Token (first 50 chars): {token[:50]}...")
    print("-" * 80)
    
    for endpoint, params in endpoints:
        try:
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            
            endpoint_desc = f"{endpoint}"
            if params:
                endpoint_desc += f" (api-version={params.get('api-version', 'none')})"
            
            print(f"\nTrying: {endpoint_desc}")
            response = requests.get(endpoint, headers=headers, params=params)
            
            print(f"Status Code: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                print(f"\n✓ Success! Response:")
                print(json.dumps(data, indent=2))
                
                # If data contains assistants/agents, list them
                if isinstance(data, dict):
                    if 'data' in data:
                        assistants = data['data']
                        print(f"\n\n{'='*80}")
                        print(f"Found {len(assistants)} assistant(s):")
                        print(f"{'='*80}")
                        for i, assistant in enumerate(assistants, 1):
                            print(f"\n{i}. {assistant.get('name', assistant.get('id', 'Unknown'))}")
                            print(f"   ID: {assistant.get('id')}")
                            print(f"   Model: {assistant.get('model', 'N/A')}")
                            if 'description' in assistant:
                                print(f"   Description: {assistant.get('description')}")
                            if 'created_at' in assistant:
                                print(f"   Created: {assistant.get('created_at')}")
                    elif 'value' in data:
                        assistants = data['value']
                        print(f"\n\n{'='*80}")
                        print(f"Found {len(assistants)} assistant(s):")
                        print(f"{'='*80}")
                        for i, assistant in enumerate(assistants, 1):
                            print(f"\n{i}. {assistant.get('name', assistant.get('id', 'Unknown'))}")
                            print(f"   ID: {assistant.get('id')}")
                elif isinstance(data, list):
                    print(f"\n\n{'='*80}")
                    print(f"Found {len(data)} assistant(s):")
                    print(f"{'='*80}")
                    for i, assistant in enumerate(data, 1):
                        print(f"\n{i}. {assistant.get('name', assistant.get('id', 'Unknown'))}")
                        print(f"   ID: {assistant.get('id')}")
                
                return  # Success, exit function
                
            elif response.status_code == 401:
                print(f"✗ Authentication failed")
                try:
                    error_data = response.json()
                    print(f"  Error: {json.dumps(error_data, indent=2)}")
                except:
                    print(f"  Response: {response.text[:300]}")
            elif response.status_code == 403:
                print(f"✗ Forbidden (Permission Denied)")
                try:
                    error_data = response.json()
                    print(f"  Error: {json.dumps(error_data, indent=2)}")
                except:
                    print(f"  Response: {response.text[:300]}")
            elif response.status_code == 404:
                print(f"✗ Endpoint not found")
            else:
                print(f"✗ Error (Status {response.status_code})")
                try:
                    error_data = response.json()
                    print(f"  Error: {json.dumps(error_data, indent=2)}")
                except:
                    print(f"  Response: {response.text[:300]}")
                
        except Exception as e:
            print(f"✗ Exception: {str(e)}")
    
    print("\n" + "=" * 80)
    print("TROUBLESHOOTING:")
    print("1. Token Scope: Ensure token is acquired with scope 'https://ai.azure.com/.default'")
    print("2. Permissions: Grant service principal the role in Azure AI Foundry project")
    print("3. Required Action: Microsoft.CognitiveServices/accounts/AIServices/agents/read")
    print("4. Suggested Roles: Cognitive Services User, Cognitive Services Contributor")

if __name__ == "__main__":
    list_assistants()

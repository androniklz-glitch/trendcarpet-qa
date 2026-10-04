import sys
import os
import json
import secrets
import hashlib
import datetime

def get_users_file():
    return os.environ.get("QA_HUB_USERS_FILE", r"C:\Users\AndronikLindgren\.secrets\qa-hub-users.json")

def load_users():
    ufile = get_users_file()
    if not os.path.exists(ufile):
        return {"enforce": False, "users": []}
    with open(ufile, "r", encoding="utf-8") as f:
        return json.load(f)

def save_users(data):
    ufile = get_users_file()
    os.makedirs(os.path.dirname(ufile), exist_ok=True)
    with open(ufile, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def main():
    if len(sys.argv) < 2:
        print("Usage: qa_hub_users.py <add|list|remove|enforce|set-org> ...")
        sys.exit(1)
        
    cmd = sys.argv[1].lower()
    data = load_users()
    
    if cmd == "add":
        if len(sys.argv) < 4:
            print("Usage: qa_hub_users.py add <name> <role> [--org reforma|cama]")
            sys.exit(1)
        name = sys.argv[2]
        role = sys.argv[3]
        
        org = None
        args = sys.argv[4:]
        i = 0
        while i < len(args):
            arg = args[i]
            if arg == "--org" and i + 1 < len(args):
                org = args[i+1].lower()
                i += 2
            elif arg.startswith("--org="):
                org = arg.split("=", 1)[1].lower()
                i += 1
            else:
                i += 1

        if role != "admin":
            if not org or org not in ("reforma", "cama"):
                print("Error: --org (reforma or cama) is required for non-admin roles", file=sys.stderr)
                sys.exit(1)
        else:
            if org and org not in ("reforma", "cama"):
                print("Error: org must be reforma or cama", file=sys.stderr)
                sys.exit(1)

        token = secrets.token_urlsafe(24)
        token_hash = hashlib.sha256(token.encode('utf-8')).hexdigest()
        
        # Remove existing if any
        data["users"] = [u for u in data.get("users", []) if u.get("name") != name]
        
        new_user = {
            "name": name,
            "role": role,
            "token_sha256": token_hash,
            "created": datetime.datetime.now().isoformat()
        }
        if org:
            new_user["org"] = org
        data.setdefault("users", []).append(new_user)
        save_users(data)
        
        ufile = get_users_file()
        links_dir = os.path.join(os.path.dirname(ufile), "qa-hub-links")
        os.makedirs(links_dir, exist_ok=True)
        link_path = os.path.join(links_dir, f"{name}.txt")
        link_url = f"https://trendcarpet-qa.vercel.app/login?t={token}"
        
        with open(link_path, "w", encoding="utf-8") as f:
            f.write(link_url)
            
        print(link_path)
        
    elif cmd == "set-org":
        if len(sys.argv) < 4:
            print("Usage: qa_hub_users.py set-org <name> <reforma|cama>", file=sys.stderr)
            sys.exit(1)
        name = sys.argv[2]
        org = sys.argv[3].lower()
        if org not in ("reforma", "cama"):
            print("Error: org must be reforma or cama", file=sys.stderr)
            sys.exit(1)
        user = next((u for u in data.get("users", []) if u.get("name") == name), None)
        if not user:
            print(f"Error: user {name} not found", file=sys.stderr)
            sys.exit(1)
        user["org"] = org
        save_users(data)
        print(f"Set org for {name} to {org}")

    elif cmd == "list":
        for u in data.get("users", []):
            org_str = f", org={u.get('org')}" if u.get('org') else ""
            print(f"{u.get('name')} ({u.get('role')}{org_str}) - {u.get('created')}")
            
    elif cmd == "remove":
        if len(sys.argv) < 3:
            print("Usage: qa_hub_users.py remove <name>")
            sys.exit(1)
        name = sys.argv[2]
        data["users"] = [u for u in data.get("users", []) if u.get("name") != name]
        save_users(data)
        print(f"Removed {name}")
        
    elif cmd == "enforce":
        if len(sys.argv) < 3:
            print("Usage: qa_hub_users.py enforce on|off")
            sys.exit(1)
        val = sys.argv[2].lower()
        if val == "on":
            data["enforce"] = True
        elif val == "off":
            data["enforce"] = False
        save_users(data)
        print(f"Enforce set to {data['enforce']}")
        
    else:
        print("Unknown command")
        sys.exit(1)

if __name__ == "__main__":
    main()

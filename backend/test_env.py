from decouple import config
import os

print("Testing environment variable loading:")
print(f"Using config(): {config('GOOGLE_API_KEY', default='NOT_FOUND')}")
print(f"Using os.getenv(): {os.getenv('GOOGLE_API_KEY', 'NOT_FOUND')}")

# Check if .env file exists
import os.path
env_exists = os.path.exists('.env')
print(f".env file exists: {env_exists}")

if env_exists:
    with open('.env', 'r') as f:
        lines = f.readlines()
        for i, line in enumerate(lines, 1):
            if 'GOOGLE_API_KEY' in line:
                print(f"Line {i}: {repr(line)}")  # repr shows whitespace
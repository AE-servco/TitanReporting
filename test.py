import os
from google.cloud import secretmanager
from google.api_core.exceptions import AlreadyExists

# Set your GCP project ID
PROJECT_ID = "prestigious-gcp"  # replace with your project

# Path to your .env file
ENV_FILE = ".env"

def load_env_file(env_path):
    """Load key=value pairs from a .env file."""
    env_vars = {}
    with open(env_path, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            value = value.strip('"').strip("'")  # remove surrounding quotes
            env_vars[key] = value
    return env_vars

def main():
    client = secretmanager.SecretManagerServiceClient()
    parent = f"projects/{PROJECT_ID}"
    env_vars = load_env_file(ENV_FILE)

    for key, value in env_vars.items():
        secret_id = key.lower()  # GCP secret names are lowercase
        secret_path = f"{parent}/secrets/{secret_id}"

        # Create secret if it doesn't exist
        try:
            client.create_secret(
                request={
                    "parent": parent,
                    "secret_id": secret_id,
                    "secret": {"replication": {"automatic": {}}},
                }
            )
            print(f"Created secret: {secret_id}")
        except AlreadyExists:
            print(f"Secret already exists: {secret_id}")

        # Add a new version with the value
        client.add_secret_version(
            request={
                "parent": secret_path,
                "payload": {"data": value.encode("UTF-8")},
            }
        )
        print(f"Added new version for secret: {secret_id}")

if __name__ == "__main__":
    main()

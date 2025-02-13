# Deployment with Kamal

This folder contains the configuration and scripts for deploying using [Kamal](https://kamal-deploy.org/),
a Ruby-based deployment tool.

Below are some useful commands for deploying and managing the app.

See the [Kamal documentation](https://kamal-deploy.org/docs/commands) for more commands.

# Setup

## 1. Setup the development environment for deploy

### 1Password CLI

See https://developer.1password.com/docs/cli/get-started/

Note: Do not use Flatpack or snap to install 1password CLI as these do not work with the SSH agent.

You will also need to update the 1Password configuration to allow it to access the SSH key:

_~/.config/1Password/ssh/agent.toml_

```toml
[[ssh-keys]]
vault = "GSO: Open Chat Studio Team (OCS)"
```

See https://developer.1password.com/docs/ssh/agent for more details.

To test that this is working you can run:

```bash
ssh ocs@107.20.181.165
```

### AWS CLI

```bash
aws configure sso --profile ocs-misc
aws sso login --profile ocs-misc
```

Note: If you used a different profile name you will need to set the `AWS_PROFILE` environment variable to the profile name.

## 2. Setup the EC2 instance

1. Create EC2 instance
2. Attach Static IP
3. Initial instance setup

    ```shell
    sudo apt update
    sudo apt upgrade -y
    sudo apt install -y docker.io curl git
    sudo usermod -a -G docker app
    sudo adduser ocs
    ```

4. Add deploy key to 'ocs' user

## 3. Setup the domain

Create a domain / subdomain and point it to the EC2 instance static IP or DNS.

## 4. Deploy services

1. Update the `deploy/config/.kamal/secrets` file with the necessary secrets.
2. Update the `deploy/config/deploy.yml` file with the EC2 instance IP and the domain name
3. Run `kamal setup`
   
   ```shell
   cd deploy
   kamal setup
   ```

## 5. Import data into the database

```shell
docker exec -it espen-sql-api-postgres bash
$> export SOURCE_DB_URL="postgres://..."
$> pg_dump $SOURCE_DB_URL -c | psql -u espen
```

## 6. Check the services

```shell
curl https://espen-sql-api.openchatstudio.com/up
curl https://espen-sql-api.openchatstudio.com/openapi.json
```

# Useful commands

## Deploy

Deploys can be done by running:

```
kamal deploy
```

This will deploy the latest local copy of your application code.

## Accessing logs

You can view the logs using Kamal by running:

```bash
kamal app logs
```

See `kamal app logs --help` for more details.

You can also access them directly on your server using `docker logs`.

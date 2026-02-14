# HPE Support Bot

## English
The HPE Support Bot is designed to assist users in managing HPE support tasks effectively.

### Setup
1. Clone the repository.
   ```bash
   git clone https://github.com/koespruyt/HPE_Support_BOT.git
   cd HPE_Support_BOT
   ```
2. Install the necessary dependencies.
   ```bash
   pip install -r requirements.txt
   ```

### Usage
To run the bot, use the following command:
```bash
python bot.py
```

### Parameters
- `--config path_to_config` : Path to the configuration file.
- `--log_level` : Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL).

### Scheduled Tasks
The bot can also run scheduled tasks:
```bash
* * * * * /usr/bin/python /path/to/bot.py --run_task
```

### Nagios Monitoring
Configure Nagios to monitor the bot as follows:
```bash
define service {
    use                 generic-service
    host_name           your_host
    service_description HPE Support Bot
    check_command       check_hpe_bot
}
```

### Troubleshooting
If you encounter issues, check the logs located at `/var/log/hpe_bot.log` for detailed error messages.

## Nederlands
De HPE Support Bot is ontworpen om gebruikers te helpen bij het effectief beheren van HPE-ondersteuningstaken.

### Installatie
1. Clone de repository.
   ```bash
   git clone https://github.com/koespruyt/HPE_Support_BOT.git
   cd HPE_Support_BOT
   ```
2. Installeer de nodige afhankelijkheden.
   ```bash
   pip install -r requirements.txt
   ```

### Gebruik
Gebruik de volgende opdracht om de bot uit te voeren:
```bash
python bot.py
```

### Parameters
- `--config pad_naar_config` : Pad naar het configuratiebestand.
- `--log_level` : Stel het niveau van logging in (DEBUG, INFO, WARNING, ERROR, CRITICAL).

### Geplande Taken
De bot kan ook geplande taken uitvoeren:
```bash
* * * * * /usr/bin/python /pad/naar/bot.py --run_task
```

### Nagios Monitoring
Configureer Nagios om de bot te monitoren:
```bash
define service {
    use                 generic-service
    host_name           uw_host
    service_description HPE Support Bot
    check_command       check_hpe_bot
}
```

### Probleemoplossing
Als u problemen ondervindt, controleer dan de logboeken op `/var/log/hpe_bot.log` voor gedetailleerde foutmeldingen.

## Security Notes
- Always keep your dependencies updated.
- Monitor the bot’s behavior for any suspicious activities.

## Repository Hygiene
- Keep your README.md file up to date.
- Regularly review and tidy up the codebase as necessary.
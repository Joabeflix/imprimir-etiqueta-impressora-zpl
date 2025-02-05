import json

try:
    with open(file='settings.json', mode='r', encoding='utf8') as file:
        config = json.load(file)
except FileNotFoundError:
    print('Erro ao ler o arquivo de configuração.')



TAB_NAME=config['tab_name']
PRINTERS_TO_IGNORE=config['printers_to_ignore']







    






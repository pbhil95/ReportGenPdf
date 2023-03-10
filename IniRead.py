import configparser

def ConfigFile():
    config = configparser.ConfigParser()
    config.sections()
    config.read('GlobalVariables.ini')
    config.sections()

    ConfigFile.Default= config['DEFAULT']
    #print(ConfigFile.Default['ResultFile'])

ConfigFile()

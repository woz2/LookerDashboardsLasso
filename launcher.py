import sys
from controller.main import gds_automation
from controller.cmdmenu import menu
from os import listdir
from os.path import isfile, join

configspath = "configs"
baseconfig = "base_config.xlsx"

configfiles = [f for f in listdir(configspath) if isfile(join(configspath, f))]
configfiles.remove(baseconfig)
configfiles = ["/configs/"+f for f in configfiles]
jobconfig = ''

if len(sys.argv) > 1:
    print(f"\nargv 1: {sys.argv[1]}")
    if sys.argv[1] in configfiles:
        try:
            jobconfig = sys.argv[1]
            print(jobconfig)
            gds_automation(jobconfig)
        except:
            print("Didn't work")
            pass
    else:
        print("Specified file does not exist in configs folder. Check file name.")
else:
    jobconfig = menu(configfiles)
    gds_automation(jobconfig)

import os
import shutil

outlook = R'M:\OUTLOOK'

attmnts = [f for f in os.scandir(outlook) if f.is_file()]
directories = [root for root, dirs, filenames in os.walk(outlook) if root != outlook]
print(len(attmnts))
for dirs in directories:
    for files in os.scandir(dirs):
        for app in attmnts:
            if app.name == files.name:
                try:
                    os.remove(app.path)
                    print(f'removed {app.name}')
                except (FileNotFoundError, PermissionError) as err:
                    #print(err)
                    pass
            elif files.name == F"{app.name[:app.name.find(' Attachment')]}.msg":
                try:
                    shutil.copyfile(app.path, RF'{dirs}\{app.name}')
                    os.remove(app.path)
                    print(f'moved {app.name} to {dirs[44:]}')
                except (FileNotFoundError, PermissionError) as err:
                    #print(err)
                    pass
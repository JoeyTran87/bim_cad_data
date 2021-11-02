import subprocess
from sys import stdout

# run = subprocess.run('dir',shell= True)
# run = subprocess.run('dir',shell= True,capture_output= True)
# run = subprocess.run('dir',shell= True,capture_output= True,text=True)
run = subprocess.run('dir',shell= True,stdout= subprocess.PIPE,text=True) # tdout= subprocess.PIPE ===> capture_output= True
print (run.returncode) #check return error
print (run.args) # check Input argrument
print (run.stdout) # check OUTPUT , need capture_output is True

# print(run.stdout.decode()) # no need if text is True


import os
import subprocess
#from pkg_resources import resource_filename

# Create conftest symlink if needed
#confpath = resource_filename(__name__, 'suppl/conftest.py')
#os.system("rm -f " + os.getcwd() + '/conftest.py')
#os.system("ln -s " + confpath)

# This is a place for steps just prior to PUTS qualification to happen.
# conftest.py symlink step removed and now happens upon invocation of virtual environment via sitecustomize.py
import subprocess

scripts_to_run = [
    'StudentSort.py',
    'ProjectSort.py',
    'ProjectAloting.py',
    'FillingNAs.py',
    'ProjectAlotingN.py',
    'FillingNAsN.py'
]

for script in scripts_to_run:
    subprocess.run(['python', script])

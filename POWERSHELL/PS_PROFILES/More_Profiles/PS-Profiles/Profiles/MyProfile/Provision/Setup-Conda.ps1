#Requires -RunAsAdministrator

choco install -s=chocolatey miniconda3 --params="'/AddToPath:1'" -y
conda install ipython jupyter -y

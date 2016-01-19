Create the Volumes and Control files by using the templates.
Any values that don't exist, leave the cells blank.

# USAGE
# Let's say you have your files in a folder named "test"
cd {go to "test" directory}
# If you have control, the command is:
ruby ms_autoformat.rb test/ test/VOLUMES.xlsx CONTROL.xlsx "creatinine"
# If you don't have control, the command is:
ruby ms_autoformat.rb test/ test/VOLUMES.xlsx



## Alternative: Standalone VBA module
Sometimes, it might be useful to run xlwings code without having to install an add-in first. To do so, you need to use the standalone option when creating a new project: 
'''
xlwings quickstart myproject --standalone.
'''

This will add the content of the add-in as a single VBA module so you don’t need to set a reference to the add-in anymore. It will also include Dictionary.cls as this is required on macOS. It will still read in the settings from your xlwings.conf if you don’t override them by using a sheet with the name xlwings.conf.
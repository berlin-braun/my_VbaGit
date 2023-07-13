# my_VbaGit
Clones a git-repository to a local working-directory and imports object into the current Acccess database

# Prerequsites
- git installed on local desktop https://git-scm.com/downloads
- a repository with Access objects filled via Application.SaveToText.

# Setup
Download class 'my_Git_Object' and import via Application.LoadfromText into a new or existing Access database.
The module 'mdl_Git' contains factory-functions for 'my_Git_Object'.

# Example
The module 'mdl_Git' contains example calls for cloning and importing several repositories.


# Report

## Quick introduce
Project that creates production report Excel files, copies some data from one Excel file to another,
fills cells with color and so on. This script has been implemented in the production plant and saves
about 2-3 man-hours for an office worker per day.

The script is working with Excel input files that contain too sensitive information to be publicly
available, but in folder 'test_data' you can find some demo versions of input files.
In testing mode 

<details>
<summary> <b>How prepare and run script?</b></summary>

1. Clone this project
2. You need to have installed Python 3 (script was developed on version 3.10)
3. Prepare environments and install requirements by typing in command line:
- go to folder where you cloned project from repository
~~~Windows PowerShell
PS> cd "path_with_cloned_project"
~~~
- create virtual environment
~~~Windows PowerShell
PS> python -m venv venv
~~~
- activate it
~~~Windows PowerShell
PS> venv\Scripts\activate
~~~
- ensure you are using virtual environment what you can check in text in your console (venv) and
install requirements
~~~Windows PowerShell
(venv) PS> python -m pip install -r requirements.txt
~~~
4. asd

</details>
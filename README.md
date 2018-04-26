# Init Environ

 - Install [`Python 3.5.2 x64`](https://www.python.org/ftp/python/3.5.2/python-3.5.2-amd64.exe)
     (run as administrator)
 - Install [`Microsoft Visual C++ Build Tools 2015`](http://landinghub.visualstudio.com/visual-cpp-build-tools)
 - Install [`virtualenvwrapper-win`](https://github.com/davidmarble/virtualenvwrapper-win/)

    ```bash
    pip install virtualenvwrapper-win
    ```
 - Create `nts` virtual environment

    ```bash
    mkvirtualenv nts --python=path-to-python35
    ```
 - Activate it

    ```bash
    workon nts
    ```
 - At this step `pip freeze` must show no installed libraries.
 - Install the required libraries:

    ```bash
    pip install -r requirements.dev
    ```

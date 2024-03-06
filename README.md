# Private_Review_Report_VNU_SIS
This application is supporting for VNU_SIS to verify their information

## Instructions
As always ensure you create a virtual environment for this application and install
the necessary libraries from the `requirements.txt` file.

```
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

Start the development server

```
$ python run.py
```


Browse to http://0.0.0.0:8080

## Docker
- sudo docker-compose up -d --build

## Guide to run
- /users: list users that imported from excel file (.xls, .xlsx)
- /import_excel: import file excel and update users to database
- /export_excel: export file excel that has UUID, Name and message to each user
```
### Trong Binh Hoang - LQDTU

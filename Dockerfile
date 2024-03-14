FROM python:3.11.5-windowsservercore-ltsc2022

ENV TZ="America/Los_Angeles"
ENV PYTHONUNBUFFERED=1

# set a directory for the app
WORKDIR /usr/src/app

# copy all the files to the container
COPY . .


RUN python -m pip install --upgrade pip
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# define the port number the container should expose
EXPOSE 5000

# run the command
CMD ["python", "./main.py"]
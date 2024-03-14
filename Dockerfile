FROM python:3.11.7-windowsservercore-ltsc2022

ENV TZ="America/Los_Angeles"
ENV PYTHONUNBUFFERED=1

# set a directory for the app
WORKDIR /app

# copy all the files to the container
COPY . .

RUN python -m pip install --upgrade pip
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# define the port number the container should expose
EXPOSE 4998

# run the command
CMD ["waitress-serve", "--host", "0.0.0.0", "--port", "4998", "main:flask_app"]

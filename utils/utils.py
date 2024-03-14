from datetime import datetime

def timestamp_now(compact=False, only_ymd=False) -> str: 
    """Credit: Brandon
        Returns a string of the current date+time in the form of
        YYYY-MM-DD hh:mm:ss
    If `compact` == True, then returns in the form of
        YYYYMMDD_hhmmss
    If `only_ymd` == True, then only the first "year/month/day" portion is returned:
        YYYY-MM-DD or YYYYMMDD
    """
    timestamp = datetime.now()
    if compact:
        if only_ymd:
            return timestamp.strftime("%Y%m%d")
        return timestamp.strftime("%Y%m%d_%H%M%S")
    if only_ymd:
        return timestamp.strftime("%Y-%m-%d")
    return timestamp.strftime("%Y-%m-%d %H:%M:%S")

def get_connection_str(filename, type):
    with open(filename, 'r') as file:
        connection_str = file.read().rstrip().split(";")

        if type == 1:
            connection_str.pop(0)
            return f'"{";".join(connection_str)[1:-1]}'
        else:
            return f'{";".join(connection_str)[1:-1]}'

def print_log(message: str) -> None:
    print(f'[{timestamp_now()}] {message}')
    return
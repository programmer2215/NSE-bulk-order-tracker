import json

def log(ping, parse_time, update_time):
    with open("log.json", 'a') as f:
        data = {
            "ping" : ping,
            "parseTime" : parse_time,
            "updateTime" : update_time
        }
        json.dump(data, f)
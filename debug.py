import re

def timestampToSeconds(timestamp):
    # convert timestamp in str to sec int
    match = re.match(r'(\d+):(\d+)', timestamp)

    if match:
        minutes, seconds = map(int, match.groups())
        return (minutes * 60 + seconds)
    return 0


timestamp = "(1:45 - 3:45)"

convertion = timestampToSeconds(timestamp)

print(f"\nGiven string: {timestamp}\n")
print(f"Convertion: {convertion}\n")

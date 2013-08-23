import json

f = open("experimenting/mock.json")
reps = json.loads(f.read())
print reps
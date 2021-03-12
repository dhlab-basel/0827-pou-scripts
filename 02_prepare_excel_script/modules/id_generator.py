# hash library
import hashlib
# random library
import random


def generate_id(key):
    return "pou_" + str(hashlib.sha256(key.encode('utf-8')).hexdigest())[:16]


def generate_random_id():
    unique_key = random.randint(1000000, 9999999)

    return "pou_" + str(hashlib.sha256(str(unique_key).encode('utf-8')).hexdigest())[:16]

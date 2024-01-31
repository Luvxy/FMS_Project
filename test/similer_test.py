from difflib import SequenceMatcher 
def similar(a, b):    
    return SequenceMatcher(None, a, b).ratio()

a = "기타)한우소머리국밥"
b = "글로벌조경)한우소머리국밥"

print(similar(a,b))
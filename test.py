import re

point_name = "TEST_SUB1_PPC_PooPoo pee pfds RTT"
if point_name.count('_') == 3:
    attribute_name = re.sub(".*_.*_.*_(.*)", r"\1", point_name)
    attribute_words = attribute_name.split(" ")
    print(attribute_name)
    print(attribute_words)

print('0'.isupper())
print('0'.islower())
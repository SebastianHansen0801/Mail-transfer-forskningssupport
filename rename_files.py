import os

PATH = os.path.join('C:\\', 'Users', 'solha', 'Desktop', 'folder1')

# i = 1
# for entry in os.listdir(PATH):
#     if os.path.isfile(os.path.join(PATH, entry)):
#         new_name = str(i) + entry
#         os.rename(os.path.join(PATH, entry), os.path.join(PATH, new_name))
#         i += 1

print(os.listdir(PATH))
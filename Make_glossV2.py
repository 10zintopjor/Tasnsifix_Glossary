import pickle
itemlist = ["as", "asdas", "asdasd"]
itemslist = ["zzas", "zzasdas", "zzasdasd"]
with open('outfile.txt', 'w') as fp:
    for item in itemlist:
        fp.write(item)
        fp.write(",")
    fp.write()

with open('outfile','rb') as fp:
    print(pickle.load(fp))
    print(pickle.load(fp))

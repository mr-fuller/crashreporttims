def txtCleanUp(txt):
    with open(txt, "r+") as f:
        line = next(f)
        f.seek(0)
        print(len(line))
        print(line[len(line) - 2])
        if line[len(line) - 2] != ',':
            f.write(line.replace("\n", ",\n"))
        else:
            pass


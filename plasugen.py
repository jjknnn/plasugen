import random
from openpyxl import load_workbook, Workbook


def open_file():
    names = {}
    wb = load_workbook(filename="nimilista.xlsx")

    sheet = wb.active

    for count, column in enumerate(sheet.iter_cols(values_only=True)):
        none_removed_list = []
        for val in list(column):
            if val is not None:
                none_removed_list.append(val)
        names[count] = none_removed_list
    return names


def ask_user_tables():
    ans = True
    i = 0
    print("Määritetään pöytien määrät ja koot")
    tables = {}
    while ans:
        places = input("Syötä pöydän paikkamäärä: ")
        amount = input("Syötä kyseisten pöytien määrä: ")
        a = 0
        while a < int(amount):
            tables[i] = places
            a += 1
            i += 1
        i += 1
        cont = input("Jatketaanko? y/n ")
        if cont != "y":
            ans = False
    return tables


def ask_user_type():
    print("Miten haluat pöydät järjestettävän: ")
    inp = input("R(andom) / J(ärjestys): ")
    if inp.lower == "r" or "random":
        return True
    return False


def fill_names_random(names, tables):
    tables_with_names = {}
    all_names = []
    for name in names:
        all_names.extend(names[name])
    random.shuffle(all_names)
    for i, table in enumerate(tables):
        tables_with_names[i] = make_tuples(all_names[: int(tables[table])])
        all_names = all_names[int(tables[table]) :]
    return tables_with_names


def make_tuples(input_list):
    result = []
    for i in range(0, len(input_list), 2):
        result.append((input_list[i], input_list[i + 1]))
    return result


def make_file(generated):
    wb = Workbook()
    ws = wb.active
    mem = 1
    for names in generated:
        for name in generated[names]:
            if len(name) != 1:
                print(name)
                ws[f"A{mem}"], ws[f"B{mem}"] = name
                mem += 1
            else:
                ws[ws[f"A{mem}"]] = name
        ws[f"A{mem}"], ws[f"B{mem}"] = ("", "")
        mem += 1
    wb.save("names.xlsx")


def balance_table(names):
    completed_name_list = []
    list_list = []
    for list in names:
        list_list.append(names[list])
    for j, name_list in enumerate(list_list, 1):
        if j < len(list_list):
            if len(list_list[j]) < len(list_list[j - 1]):
                list1 = list_list[j - 1]
                list2 = list_list[j]
            else:
                list2 = list_list[j - 1]
                list1 = list_list[j]
        else:
            break
        len1 = len(list1)
        len2 = len(list2)
        total_length = len1 + len2
        repetitions = total_length // len2
        merged_list = []
        for i in range(max(len1, len2)):
            if i < len1:
                merged_list.append(list1[i])
            if i % repetitions == 0 and i // repetitions < len2:
                merged_list.append(list2[i // repetitions])
        list_list[j] = merged_list
        completed_name_list = merged_list
    no_dups = []
    for item in completed_name_list:
        if item not in no_dups:
            no_dups.append(item)
    for name in names:
        for n in names[name]:
            if n not in no_dups:
                random_index = random.randint(0, len(no_dups))
                no_dups.insert(random_index, n)
    print(no_dups)
    print(len(no_dups))

    return no_dups


def get_weights(names):
    weights = []
    total = 0
    for name in names:
        total += len(names[name])
    for name in names:
        if len(names[name]) != 0:
            weights.append(len(names[name]) / total)
    return weights


def fill_names_order(names, tables):
    for name in names:
        print(name)
    balanced = balance_table(names)
    print(len(balanced))
    tables_with_names = {}
    for i, table in enumerate(tables):
        tables_with_names[i] = make_tuples(balanced[: int(tables[table])])
        balanced = balanced[int(tables[table]) :]
    return tables_with_names


def make_file_order(generated):
    wb = Workbook()
    ws = wb.active
    mem = 1
    for names in generated:
        for name in generated[names]:
            print(name)
            ws[f"A{mem}"], ws[f"B{mem}"] = name
            mem += 1
        ws[f"A{mem}"], ws[f"B{mem}"] = ("", "")
        mem += 1
    wb.save("names.xlsx")


def shuffle_names(names):
    shuffled = {}
    for name in names:
        if len(names[name]) != 0:
            random.shuffle(names[name])
            shuffled[name] = names[name]
    return shuffled


def main():
    names = open_file()
    names = shuffle_names(names)
    random = ask_user_type()
    tables = ask_user_tables()
    if random:
        generated = fill_names_random(names, tables)
    else:
        generated = fill_names_order(names, tables)
    make_file(generated)


main()

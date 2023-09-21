import random
from openpyxl import load_workbook, Workbook


def open_file():
    """
    Read xlsx file
    """
    names = {}
    wb = load_workbook(filename="nimilista.xlsx")

    sheet = wb.active

    # Fetch names from xlsx based on columns
    for count, column in enumerate(sheet.iter_cols(values_only=True)):
        none_removed_list = []
        # Remove empty cells
        for val in list(column):
            if val is not None:
                none_removed_list.append(val)
        names[count] = none_removed_list
    return names


def ask_user_tables(count):
    """
    Ask user for table sizes and amounts
    """
    ans = True
    i = 0
    print("\nMääritetään pöytien määrät ja koot")
    tables = {}
    while ans:
        print(f"Nimiä jäljellä: {count}")
        places = input("Syötä pöydän paikkamäärä: ")
        amount = input("Syötä kyseisten pöytien määrä: ")
        if int(places)*int(amount) > count:
            print("\nSyötit enemmän paikkoja kun listassa on nimiä!\nPöytää ei lisätty!")
            continue
        a = 0
        while a < int(amount):
            tables[i] = places
            a += 1
            i += 1
        count -= int(places)*int(amount)
        if count == 0:
            print("Kaikki paikat täytetty!")
            break
        i += 1
        cont = input("Jatketaanko? y/n ")
        if cont != "y":
            ans = False
        print()  # Newline to help readability
    return tables


def ask_user_type():
    """
    Ask user for name sorting type
    """
    print("Miten haluat pöydät järjestettävän: ")
    inp = input("R(andom) / J(ärjestys): ")
    if inp.lower == "r" or "random":
        return True
    return False


def fill_names_random(names, tables):
    """
    Fill names to given tables randomly
    """
    tables_with_names = {}
    all_names = []
    for name in names:
        all_names.extend(names[name])

    # Shuffle names
    random.shuffle(all_names)

    # Generate tables with names
    for i, table in enumerate(tables):
        tables_with_names[i] = make_tuples(all_names[: int(tables[table])])
        all_names = all_names[int(tables[table]) :]
    return tables_with_names


def make_tuples(input_list):
    """
    Makes tuples from the names for feeding to tables
    """
    result = []
    if len(input_list) % 2 == 0:
        for i in range(0, len(input_list), 2):
            result.append((input_list[i], input_list[i + 1]))
    else:
        for i in range(0, len(input_list) - 1, 2):
            result.append((input_list[i], input_list[i + 1]))
        result.append((input_list[i+2], ""))
    return result


def make_file(generated):
    """
    Make xlsx file with A and B column with names
    """
    wb = Workbook()
    ws = wb.active
    mem = 1
    for names in generated:
        for name in generated[names]:
            if len(name) >= 2:
                ws[f"A{mem}"], ws[f"B{mem}"] = name
                mem += 1
            elif len(name) == 1:
                ws[ws[f"A{mem}"]] = name
            else:
                continue
        ws[f"A{mem}"], ws[f"B{mem}"] = ("", "")
        mem += 1
    wb.save("names.xlsx")


def balance_table(names):
    """
    Generate names to such order that they are evenly divided to the tables
    TODO: Reformat this method
    """
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
    return no_dups


def fill_names_order(names, tables):
    """
    Fill names to tables when using ordered method
    """
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
    """
    Makes the xlsx file when using ordered method
    """
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
    """
    Shuffles names from xlsx
    """
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
    amount = 0
    for name in names:
        amount += len(names[name])
    tables = ask_user_tables(amount)
    if random:
        generated = fill_names_random(names, tables)
    else:
        generated = fill_names_order(names, tables)
    make_file(generated)
    print("names.xlsx luotu! Katso kansion sisältö!")


main()

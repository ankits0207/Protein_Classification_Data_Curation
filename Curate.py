from xlrd import open_workbook
import csv


class MyClass:
    def __init__(self, org_acronym, pdb_id, structure_title, resolution, structure_mw, macromolecule_type, residue_count):
        self.org_acronym = org_acronym
        self.pdb_id = pdb_id
        self.structure_title = structure_title
        self.resolution = resolution
        self.structure_mw = structure_mw
        self.macromolecule_type = macromolecule_type
        self.residue_count = residue_count


def check_str_in_list(my_l, my_str):
    for elt in my_l:
        if elt == my_str:
            return 1
    return 0


def get_no_of_max_mw_obj(my_l):
    max_mw = my_l[0].structure_mw
    for obj in my_l:
        if obj.structure_mw > max_mw:
            max_mw = obj.structure_mw
    count = 0
    for obj in my_l:
        if obj.structure_mw == max_mw:
            count += 1
    return count, max_mw


def get_no_of_max_rc_obj(my_l):
    max_rc = my_l[0].residue_count
    for obj in my_l:
        if obj.residue_count > max_rc:
            max_rc = obj.residue_count
    count = 0
    for obj in my_l:
        if obj.residue_count == max_rc:
            count += 1
    return count, max_rc


def get_no_of_min_res_obj(my_l):
    min_res = my_l[0].resolution
    for obj in my_l:
        if obj.resolution < min_res:
            min_res = obj.resolution
    count = 0
    for obj in my_l:
        if obj.resolution == min_res:
            count += 1
    return count, min_res


def get_useless_objects(my_l):
    list_to_be_returned = []
    count, max_mw = get_no_of_max_mw_obj(my_l)
    if count == 1:
        idx = -1
        for k in range(len(my_l)):
            if my_l[k].structure_mw == max_mw:
                idx = k
        for k in range(len(my_l)):
            if k != idx:
                list_to_be_returned.append(my_l[k])
        return list_to_be_returned
    count, max_rc = get_no_of_max_rc_obj(my_l)
    if count == 1:
        idx = -1
        for k in range(len(my_l)):
            if my_l[k].residue_count == max_rc:
                idx = k
        for k in range(len(my_l)):
            if k != idx:
                list_to_be_returned.append(my_l[k])
        return list_to_be_returned
    count, min_res = get_no_of_min_res_obj(my_l)
    if count == 1:
        idx = -1
        for k in range(len(my_l)):
            if my_l[k].resolution == min_res:
                idx = k
        for k in range(len(my_l)):
            if k != idx:
                list_to_be_returned.append(my_l[k])
        return list_to_be_returned
    return list_to_be_returned


list_of_objects = []
wb = open_workbook('Mesophilic.xlsx')
for sheet in wb.sheets():
    rows = sheet.nrows
    columns = sheet.ncols
    for row in range(1, rows):
        temp_list = []
        for column in range(columns):
            temp_list.append(sheet.cell(row, column).value)
        acro = temp_list[0]
        pdb = temp_list[1]
        title = temp_list[2]
        reso = temp_list[3]
        mw = temp_list[4]
        type = temp_list[5]
        res_count = temp_list[6]
        my_obj = MyClass(str(acro).strip(), str(pdb).strip(), str(title).strip(), reso, mw, str(type).strip(), res_count)
        list_of_objects.append(my_obj)

len_before = len(list_of_objects)

# Check over resolution, if resolution > 2.5, then dump that object
temp_list = []
for obj in list_of_objects:
    if obj.resolution > 2.5:
        temp_list.append(obj)
for obj in temp_list:
    list_of_objects.remove(obj)

# Keep only unique pdb ids
temp_list = []
for i in range(0, len(list_of_objects)):
    obj1 = list_of_objects[i]
    for j in range(i+1, len(list_of_objects)):
        obj2 = list_of_objects[j]
        if obj1.pdb_id == obj2.pdb_id:
            temp_list.append(obj2.pdb_id)
for pdb in temp_list:
    for obj in list_of_objects:
        if pdb == obj.pdb_id:
            list_of_objects.remove(obj)

# Keep only unique titles
my_dict = {}
list_of_str = []
for i in range(0, len(list_of_objects)):
    if check_str_in_list(list_of_str, list_of_objects[i].structure_title) == 0:
        obj1 = list_of_objects[i]
        my_list = []
        flag = 0
        for j in range(i+1, len(list_of_objects)):
            obj2 = list_of_objects[j]
            if obj1.structure_title == obj2.structure_title:
                list_of_str.append(list_of_objects[i].structure_title)
                my_list.append(obj2)
                flag = 1
        if flag == 1:
            my_list.append(obj1)
            my_dict[obj1.structure_title] = my_list
        list_of_str = list(set(list_of_str))

cleanup_list = []
for key, list_of_values in my_dict.iteritems():
    useless_objects = get_useless_objects(list_of_values)
    if len(useless_objects) != 0:
        for obj in useless_objects:
            cleanup_list.append(obj)
    else:
        for i in range(len(list_of_values) - 1):
            cleanup_list.append(list_of_values[i])

for obj1 in cleanup_list:
    cleanup_pdb = obj1.pdb_id
    for obj2 in list_of_objects:
        if obj2.pdb_id == cleanup_pdb:
            list_of_objects.remove(obj2)

len_after = len(list_of_objects)

print len_before - len_after

with open('M.csv', 'w') as csvfile:
    list_of_labels = ['PDB_ID', 'STRUCTURE_TITLE', 'RESOLUTION', 'MW', 'TYPE', 'RES_COUNT']
    writer = csv.DictWriter(csvfile, fieldnames=list_of_labels)
    writer.writeheader()
    for obj in list_of_objects:
        writer.writerow({'PDB_ID': obj.pdb_id, 'STRUCTURE_TITLE': obj.structure_title, 'RESOLUTION': obj.resolution, 'MW': obj.structure_mw, 'TYPE': obj.macromolecule_type, 'RES_COUNT': obj.residue_count})

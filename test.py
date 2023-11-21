l = ['1',2,'1',1,3,1,2,2,7]


def identical_column_names(list_to_check):
    list_to_check = list_to_check
    identical_column_names_list = []
    while list_to_check:
        i = list_to_check[0]
        list_to_check = list_to_check[1:]
        if i in list_to_check and i not in identical_column_names_list:
            identical_column_names_list.append(i)
    return identical_column_names_list

lr = identical_column_names(l)

print(lr)
print(' ,'.join(list(map(str,lr))))
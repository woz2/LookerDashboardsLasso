dic = {}


def print_menu(lst):
    for i in lst:
        dic[lst.index(i) + 1] = i
    keys = list(dic.keys())
    dic[keys[-1]+1] = "EXIT"
    for key in dic.keys():
        print(key, '--', dic[key])
    return dic


def menu(lst):
    while True:
        print("No input was provided.\nAvailable config files:")
        print_menu(lst)
        keys = list(dic.keys())
        option = ''
        try:
            option = int(input("Enter number to select a config file or chose 'EXIT' to abort: "))
            selected = dic[option]
        except:
            print('Wrong input. Please enter a number ...')
        # Check what choice was entered and act accordingly
        if option < keys[-1]:
            print(f"You have selected: {selected}")
            return selected
        elif option == keys[-1]:
            print("Function aborted!")
            exit()
        else:
            print('Invalid option. Please enter a number between 1 and 4.')


if __name__ == '__main__':
    menu()

repl = True
while(repl):
    from random import *

    print()

    while(True):
        try:
            people = int(input("인원수를 입력하시오> "))
            print()
            if people > 0:
                break
            else:
                raise ValueError
        except ValueError:
            print("\n잘못된 값을 입력하였습니다.\n")

    name = []
    prize_item = []

    for i in range(people):
        while(True):
                name.append(input("이름을 적으시오> "))
                print()
                if name[i].isspace():
                    print("공백이 입력되었습니다.")
                    name.pop()
                    print()
                elif not name[i]:
                    print("공백이 입력되었습니다.")
                    name.pop()
                    print()
                else:
                    break

        for i0 in range(people):
            try:
                print(name[i0] + " ->")
            except IndexError:
                print("->")
        print()

    for i in range(people):
        while(True):
                prize_item.append(input("항목을 적으시오> "))
                print()
                if prize_item[i].isspace():
                    print("공백이 입력되었습니다.")
                    prize_item.pop()
                    print()
                elif not prize_item[i]:
                    print("공백이 입력되었습니다.")
                    prize_item.pop()
                    print()
                else:
                    break
                
        for i0 in range(people):
            try:
                print(f"{name[i0]} -> {prize_item[i0]}")
            except IndexError:
                print(f"{name[i0]} ->")
        print()

    shuffle(name)
    shuffle(prize_item)

    print("최종 결과는?\n")
    for i in range(people):
        print(f"{name[i]} -> {prize_item[i]}")
    print()
    re = input(str("다시? (0 입력시 종료)"))
    re = re.strip()
    if re == "0":
        repl = False

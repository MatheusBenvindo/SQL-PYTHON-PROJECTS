

cpf = str(input("digite o cpf"))

print (cpf)

i = [10,9,8,7,6,5,4,3,2]
i2 = [1,2,3,4,5,6,7,8,9,0]

def verdig1 (cpf):
    sum = 0
    global i
    for digito in cpf:
        sum += (digito) * i 
        print (f"{digito} {i} {sum}")
        i -= 1

    
    resto = (sum%11)
    print (resto)

def verdig2 (cpf):
    sum = 0
    for digito in cpf:
        sum += int(digito) * i 
        print (f"{digito} {i} {sum}")
        i -= -1

    
    resto = (sum%11-6)
    print (resto)

verdig1(cpf)
verdig2(cpf)



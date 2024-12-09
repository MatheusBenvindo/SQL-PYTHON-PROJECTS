cpf = str(input("digite o cpf"))

from unidecode import unidecode


def clear_cpf(cpf):
    return cpf.replace("-", "").replace(".", "")


def validar_cpf(cpf):
    cpf = clear_cpf(cpf)
    sum = 0
    i = 10
    for digito in range(0, 9):
        sum += int(cpf[digito]) * i
        print(f"{cpf[digito]} {i} {sum}")
        i -= 1

    resto = sum % 11
    if resto < 2:
        resto = 0
    else:
        resto = 11 - resto
    print(resto)

    sum = 0
    i = 11
    for digito in range(0, 10):
        sum += int(cpf[digito]) * i
        print(f"{cpf[digito]} {i} {sum}")
        i -= 1

    resto = sum % 11
    if resto < 2:
        resto = 0
    else:
        resto = 11 - resto

    print(resto)


validar_cpf(cpf)

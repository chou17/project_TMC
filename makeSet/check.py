def checkarr(str):
    if str == '發散表寒':
        return 0
    elif str == '祛風寒':
        return 1
    else:
        return -1

    # ...(可能要寫所有的or有其他方法？)


def checkeffect(int):
    if int == 0:
        return '發散表寒'
    if int == 1:
        return '祛風寒'

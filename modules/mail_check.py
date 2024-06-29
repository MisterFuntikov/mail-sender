import re


def mailCheck(mail: str, param={}) -> dict:
    '''
    PARAM = {
        'empty_mail_ignore': False,
        'russian_letter_ignore': False,
        'intell_mail_check': False,
        'multiple_mail': False,
        'split_symbol': ''
    }
    '''
    # split_symbol ~ ';|; '

    mail = str(mail)

    if type(param) != dict:
        param = {}

    if not 'empty_mail_ignore' in param or (param['empty_mail_ignore'] != False and param['empty_mail_ignore'] != True):
        param['empty_mail_ignore'] = False
    if not 'russian_letter_ignore' in param or (param['russian_letter_ignore'] != False and param['russian_letter_ignore'] != True):
        param['russian_letter_ignore'] = False
    if not 'intell_mail_check' in param or (param['intell_mail_check'] != False and param['intell_mail_check'] != True):
        param['intell_mail_check'] = False
    if not 'multiple_mail' in param or (param['multiple_mail'] != False and param['multiple_mail'] != True):
        param['multiple_mail'] = False

    if param['multiple_mail'] == True:
        if not 'split_symbol' in param or param['split_symbol'] == '':
            masmail = [mail]
        else:
            masmail = re.split(param['split_symbol'], mail)
    else:
        masmail = [mail]

    if param['russian_letter_ignore'] == True:
        pattern = r'^[\w\-\+]+[\w\-\+\.]*\@[\w\-\+]+(\.[\w\-\+]+)+$'
    else:
        pattern = r'^[0-9A-Za-z\_\-\+]+[0-9A-Za-z\_\-\+\.]*\@[0-9A-Za-z\_\-\+]+(\.[0-9A-Za-z\_\-\+]+)+$'

    if mail == '' or mail == None or mail == 'None':
        if param['empty_mail_ignore'] == True:
            return {'status': True}
        else:
            return {'status': False, 'msg': 'Пустое значение'}

    colerr = 0
    maserr = ''

    for cmail in masmail:
        if not re.fullmatch(pattern=pattern, string=cmail):
            colerr += 1
            if colerr > 1:
                maserr += '\n'
            maserr += f'Неверный почтовый адрес "{cmail}"'

            if param['intell_mail_check'] == True:

                if re.search(r'\s', cmail):
                    maserr += '\n - Содержит пробельный символ'

                if param['russian_letter_ignore'] == False \
                        and re.search(r'[а-яА-Я]', cmail):
                    maserr += '\n - Содержит кириллицу'

                if not re.search(r'\@', cmail):
                    maserr += '\n - Отсутствует почтовый домен'
                elif re.search(r'^[\w\-\+]+[\w\-\+\.]*\@', cmail) \
                        and not re.search(r'\@\s*[\s\w\-\+]+(\.[\s\w\-\+]+)+$', cmail):
                    maserr += '\n - Неправильный почтовый домен'

    if colerr == 0:
        return {'status': True}
    return {'status': False, 'msg': maserr}

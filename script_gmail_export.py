import datetime
import email
import hashlib
import imaplib
import os
from difflib import SequenceMatcher

import pandas as pd

EMAIL_UN = 'username@domain.com'
EMAIL_PW = 'password'


def attachment_download():
    un = EMAIL_UN
    pw = EMAIL_PW
    url = 'imap.gmail.com'
    m = imaplib.IMAP4_SSL(url, 993)
    m.login(un, pw)
    m.select('Inbox')
    
    # variable to keep message from IMAP
    message_list = []
    
    # syntax date like since 19-jun-2020 to 20-jun-2020
    date = (datetime.date.today()).strftime("%d-%b-%Y")
    # before 1 days
    before_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d-%b-%Y")
    

    mail_list = m.search(
        None,
        '(SUBJECT ERROR)',
        '(SUBJECT EXTERNAL)',
        '(SUBJECT Server)',
        '(SENTSINCE %s)' % before_date,
        '(SENTBEFORE %s)' % date,
    )

    mail_list = mail_list[1][0].split()

    for _uid in range(int(mail_list[-1]), int(mail_list[0]) - 1, -1):
        email_message_dict = dict()
        count += 1
        status_subject, data_subject_list = m.fetch(str(_uid), '(RFC822.SIZE BODY[HEADER.FIELDS (SUBJECT)])')
        subject_message = data_subject_list[0][1].decode('utf-8').lstrip('Subject: ').strip() + ' '

        status, data_topic_list = m.fetch(str(_uid), '(RFC822)')
        raw_msg = data_topic_list[0][1].decode('utf-8')
        email_message_dict.update({'Subject': str(subject_message)})
        email_message_dict.update(email.message_from_string(raw_msg).__dict__)
        message_list = get_message_list(email_message_dict, message_list)
        
    df = pd.DataFrame(message_list)
    # for sort column
    # df = df[['Subject', '_headers', '_payload']]
    
    writer = pd.ExcelWriter('email_log.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='email_message')
    writer.save()


def get_message_list(dic: dict, data_list: list) -> list:
    _first_point = 0
    _last_point = 0
    _payload = None
    list_of_dic = []
    dic = lean_dict(dic)
    try:
        _payload_message = dic.get('_payload', None)
    except:
        return data_list
        
    if _payload_message and isinstance(_payload_message, list):
        for _d in _payload_message:
            # check have trackback in side?
            # traceback seem?
            _d = lean_dict(_d.__dict__)
            _payload = _d.get('_payload', None)
            if _payload and isinstance(_payload, list):
                for _p in _payload:
                    _p = (lean_dict(_p.__dict__)).get('_payload', None)
                    data_list = get_message_list(_p, data_list)
            elif _payload and 'Traceback' in _payload:
            	# hash for compare message if you got same email or reply email
                _first_point, _last_point = get_point_of_string(_payload)
                list_of_dic.append(
                    {
                        '_hash_payload': hashlib.md5(_payload[_first_point:_last_point].encode()).hexdigest(),
                        '_payload': _payload[_first_point:_last_point]
                    }
                )
            else:
                continue
                
        if len(list_of_dic) > 1:
            # Check similar
            for _index1 in range(0, len(list_of_dic)):
                for _index2 in range(_index1 + 1, len(list_of_dic)):
                    # compare 2 payload if like > 80% will continue to next round
                    if similar(list_of_dic[_index1]['_hash_payload'], list_of_dic[_index2]['_hash_payload']) >= 0.8:
                        continue
                    # compare new payload with payload list
                    if next((_dic for _dic in data_list if
                             similar(_dic['_hash_payload'], list_of_dic[_index1]['_hash_payload']) >= 0.8), None):
                        continue

                    data_list.append(
                        dict({
                            'Subject': dic['Subject'],
                            'policy': dic['policy'],
                            '_hash_payload': list_of_dic[_index1]['_hash_payload'],
                            '_payload': list_of_dic[_index1]['_payload']
                        })
                    )
            return data_list
            
    elif _payload_message and 'Traceback' in _payload_message:
        _first_point, _last_point = get_point_of_string(_payload_message)
        _hash_payload = hashlib.md5(_payload_message[_first_point:_last_point].encode()).hexdigest()
        if len(data_list) > 0 and next(
                (_dic for _dic in data_list if similar(_dic['_hash_payload'], _hash_payload) >= 0.8), None):
            return data_list
        dic.update(
            {
                'Subject': dic['Subject'],
                'policy': dic['policy'],
                '_hash_payload': _hash_payload,
                '_payload': _payload_message[_first_point:_last_point]
            }
        )
        data_list.append(dic)
        return data_list
    return data_list


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def lean_dict(dic):
    if isinstance(dic, dict):
        del_list = ['_unixfrom', '_charset', 'preamble', 'defects', 'epilogue', '_default_type', '_headers']
        for del_name in del_list:
            try:
                del dic[del_name]
            except:
                continue
    return dic

# Track only message in email that you want
def get_point_of_string(_s):
    _first_point = _s.rfind('Traceback')
    _last_point = _s.rfind('Request information')
    return _first_point, _last_point

attachment_download()


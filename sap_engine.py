import re

import customizing

class EbsEngine ():
    def __init__(
            self,
            content: str
            ):
        self.content = content

        lines = self.content.split('\n')
        self.swift = ''
        self.bank_account_number = ''
        self.start_date = ''
        self.end_date = ''
        self.currency = ''
        self.opening_balance = 0
        self.closing_balance = 0

        value_date = ''
        amount = ''
        bank_external_transaction = ''
        bank_ext_tr_description = ''
        self.position_lst = []
        self.total_debit = 0
        self.total_credit = 0

        #header and position value into variables are set here, depending on file content
        for line in lines:
            if line[:3] == '{1:':
                self.swift = line[6:14]
            elif line[:4] == ':25:':
                self.bank_account_number = line[4:]
            elif line[:5] == ':60F:':
                self.start_date = line[6:12]
                yy = self.start_date[:2]
                mm = self.start_date[2:4]
                dd = self.start_date[4:6]
                self.start_date = dd + "/" + mm + "/20" + yy
                self.currency = line[12:15]
                self.opening_balance = line[15:].replace(",", ".")
                if line[5] == 'C':
                    self.opening_balance = '-' + line[15:].replace(",", ".")
            elif line[:5] == ':62F:':
                self.end_date = line[6:12]
                yy = self.end_date[:2]
                mm = self.end_date[2:4]
                dd = self.end_date[4:6]
                self.end_date = dd + "/" + mm + "/20" + yy
                self.closing_balance = line[15:].replace(",", ".")
                if line[5] == 'C':
                    self.closing_balance = '-' + line[15:].replace(",", ".")
            elif line[:4] == ':61:':
                if line[14] == 'C' or line[14:16] == 'RD':
                    amount = "-" + re.search(r'(\d+,\d+)', line[4:]).group(1)
                    amount = "{:.2f}".format(float(amount.replace(",", ".")))
                    self.total_credit += float(amount)
                else:
                    amount = re.search(r'(\d+,\d+)', line[4:]).group(1)
                    amount = "{:.2f}".format(float(amount.replace(",", ".")))
                    self.total_debit += float(amount)
                value_date = line[4:10]
                yy = value_date[:2]
                mm = value_date[2:4]
                dd = value_date[4:6]
                value_date = dd + "/" + mm + "/20" + yy
                bank_external_transaction = line[line.find(',') + 3:line.find(',') + 7]
                if line[line.find(',') + 2].isalpha():
                    bank_external_transaction = line[line.find(',') + 2:line.find(',') + 6]
                if bank_external_transaction in customizing.ebs_mt940_dict:
                    bank_ext_tr_description = customizing.ebs_mt940_dict[bank_external_transaction]
                else:
                    bank_ext_tr_description = ''
                self.position_lst.append((value_date, amount, bank_external_transaction, bank_ext_tr_description, '', '', ''))

class IbanEngine ():
    def __init__(
            self,
            content: str
            ):
        self.content = content

        lines = self.content.split('\n')

        iban = ''
        bank_country = ''
        bank_key = ''
        bank_account_number = ''
        bank_control_key = ''
        swift = ''
        notes = ''
        self.position_lst = []
        self.iban_not_analyzed = 0

        for line in lines:
            iban = ''
            bank_country = ''
            bank_key = ''
            bank_account_number = ''
            bank_control_key = ''
            swift = ''
            notes = ''
            if line != '':
                iban = line
                bank_country = line[:2]
                if bank_country == 'IT':
                    bank_key = line[5:15]
                    bank_account_number = line[15:27]
                    bank_control_key = line[4]
                elif bank_country == 'ES':
                    bank_key = line[4:12]
                    bank_account_number = line[14:24]
                    bank_control_key = line[12:14]
                elif bank_country == 'BE':
                    bank_key = line[4:7]
                    bank_account_number = line[4:7] + '-' + line[7:14] + '-' + line[14:16]
                    notes = 'For belgian banks, it needs to enter in bank account number:- "BANK_KEY-BANK_ACCOUNT_NUMBER-BANK_CONTROL_KEY"'
                elif bank_country == 'FR':
                    bank_key = line[4:14]
                    bank_account_number = line[14:25]
                    bank_control_key = line[25:27]
                elif bank_country == 'NL':
                    swift = line[4:8]
                    bank_account_number = line[8:18]
                    notes = 'For Netherland banks, SAP bank key is not relevant to calculate IBAN, SAP extract it from SWIFT'
                elif bank_country == 'FI':
                    bank_key = line[4:10]
                    bank_account_number = line[10:17]
                    bank_control_key = line[17]
                elif bank_country == 'LU':
                    bank_key = line[4:7]
                    bank_account_number = line[7:20]
                elif bank_country == 'CH':
                    bank_key = line[4:9]
                    bank_account_number = line[9:21]
                elif bank_country == 'GB' or bank_country == 'IE':
                    swift = line[4:8]
                    bank_key = line[8:14]
                    bank_account_number = line[14:22]
                    notes = 'For UK and Ireland, it is used the SWIFT code in the IBAN. It needs to establish if it needs to include it in the SAP bank key'
                elif bank_country == 'DE':
                    bank_key = line[4:12]
                    bank_account_number = line[12:22]
                else:
                    notes = 'Error, bank country not recognized'
                    self.iban_not_analyzed += 1

                self.position_lst.append([iban, bank_country, bank_key, bank_account_number, bank_control_key, swift, notes])
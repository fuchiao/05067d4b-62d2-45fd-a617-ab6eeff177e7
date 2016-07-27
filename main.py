from openpyxl import load_workbook
import arrow

sales_list = {}
class sales(object):
    def __init__(self):
        self.job = ''
        self.fulltime = True
        self.loan = {'fyc':0, 'nfyc':0, 'nfyc_c3':0}
        self.team_income = 0
        self.supervisor_bonus = 0
        self.dt = []
        self.dt_fyc = 0

    def add_loan(self, loan, type, date):
        if arrow.get(date).replace(years=1) > arrow.now():
            self.loan['fyc'] += loan
        elif arrow.get(date).replace(years=2) > arrow.now():
            if type in ['A2', 'A3', 'B2', 'B3', 'C2']:
                self.loan['nfyc'] += loan
            elif type in ['C3']:
                self.loan['nfyc_c3'] += loan
            else:
                pass
        elif arrow.get(date).replace(years=3) > arrow.now():
            if type in ['A3', 'B3']:
                self.loan['nfyc'] += loan
            elif type in ['C3']:
                self.loan['nfyc_c3'] += loan
            else:
                pass
        else:
            pass

    def get_fyc(self):
        return self.loan['fyc'] * 0.02

    def get_nfyc(self):
        return self.loan['nfyc']

    def get_nfyc_c3(self):
        return self.loan['nfyc_c3']

    def get_direct_service_bonus(self): #收入一
        if self.job == '業務員':
            return self.loan['fyc'] * 0.02 * 0.95 * 1.5
        else:
            return self.loan['fyc'] * 0.02 * 1.5

    def get_continuous_loan_bonus(self):  # 收入二
        return (self.loan['nfyc'] * 0.01 + self.loan['nfyc_c3'] * 0.02) * 1.5

    def get_education_bonus(self):  # 收入五
        if self.job == '業務員':
            return 0;
        else:
            return self.get_member_fyc() * 0.05 * 1.5

    def get_member_fyc(self):
        sum = 0
        if self.job == '業務員':
            return 0
        for i in self.dt:
            if sales_list[i].job == '業務員':
                sum += sales_list[i].get_fyc()
        return sum

    def get_dt_fyc(self):
        return self.get_fyc() + self.get_member_fyc()

    def get_second_layer_dt_fyc(self, second_layer_job):
        sum = 0
        for i in self.dt:
            if sales_list[i].job == second_layer_job:
                sum += sales_list[i].get_dt_fyc()
        return sum

    def get_second_layer_bonus(self):
        sum = 0
        if self.job == '業務主任':
            print('Layer2 {} {}'.format('業務主任', self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.07))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.07
        elif self.job == '業務襄理':
            print('Layer2 {} {}'.format('業務主任', self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.07))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.07
            print('Layer2 {} {}'.format('業務襄理', self.get_second_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.07))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.07
        elif self.job == '業務經理':
            print('Layer2 {} {}'.format('業務主任', self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.09))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.09
            print('Layer2 {} {}'.format('業務襄理', self.get_second_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.07))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.07
            print('Layer2 {} {}'.format('業務經理', self.get_second_layer_dt_fyc(second_layer_job='業務經理') * 1.5 * 0.06))
            sum += self.get_second_layer_dt_fyc(second_layer_job='業務經理') * 1.5 * 0.06
        else:
            pass
        return sum

    def get_third_layer_dt_fyc(self, second_layer_job):
        sum = 0
        for i in self.dt:
            if sales_list[i].job == second_layer_job:
                sum += sales_list[i].get_second_layer_dt_fyc(second_layer_job='業務主任')
                sum += sales_list[i].get_second_layer_dt_fyc(second_layer_job='業務襄理')
                sum += sales_list[i].get_second_layer_dt_fyc(second_layer_job='業務經理')
        return sum

    def get_third_layer_bonus(self):
        sum = 0
        if self.job == '業務主任':
            print('Layer3 {} {}'.format('業務主任', self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.03))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.03
        elif self.job == '業務襄理':
            print('Layer3 {} {}'.format('業務襄理', self.get_third_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.03))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.03
            print('Layer3 {} {}'.format('業務主任', self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.035))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.035
        elif self.job == '業務經理':
            print('Layer3 {} {}'.format('業務經理', self.get_third_layer_dt_fyc(second_layer_job='業務經理') * 1.5 * 0.025))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務經理') * 1.5 * 0.025
            print('Layer3 {} {}'.format('業務襄理', self.get_third_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.03))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務襄理') * 1.5 * 0.03
            print('Layer3 {} {}'.format('業務主任', self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.045))
            sum += self.get_third_layer_dt_fyc(second_layer_job='業務主任') * 1.5 * 0.045
        else:
            pass
        return sum

    def get_dt_fyc_level_bonus(self):
        ret = 0
        dt_fyc = self.get_dt_fyc() * 1.5
        if self.job == '業務經理' and dt_fyc >= 120000:
            print('27%')
            ret = dt_fyc * 0.27
        elif self.job == '業務襄理' and dt_fyc >= 100000:
            print('23%')
            ret = dt_fyc * 0.23
        elif dt_fyc >= 80000:
            print('21%')
            ret = dt_fyc * 0.21
        elif dt_fyc >= 60000:
            print('16%')
            ret = dt_fyc * 0.16
        elif dt_fyc >= 45000:
            print('9%')
            ret = dt_fyc * 0.09
        elif dt_fyc >= 30000:
            print('6%')
            ret = dt_fyc * 0.06
        else:
            print('0%')
            ret = 0
        print('DTFYC LEVEL Bonus {}'.format(ret))
        return ret

    def add_direct_team_member(self, s):
        if s not in self.dt:
            self.dt.append(s)

    def get_organization_bonus(self):
        return self.get_second_layer_bonus() + self.get_third_layer_bonus() + self.get_dt_fyc_level_bonus()

def load_test():
    wb = load_workbook('test1.xlsx')
    ws = wb.active
    for row in ws.rows:
        if row[0].value == '姓名':
            continue
        if row[0].value not in sales_list:
            sales_list[row[0].value] = sales()
        sales_list[row[0].value].job = row[1].value
        if row[2].value == 'Y':
            sales_list[row[0].value].fulltime = True
        else:
            sales_list[row[0].value].fulltime = False
        if row[3].value and row[3].value not in sales_list:
            sales_list[row[3].value] = sales()
        if row[3].value in sales_list:
            sales_list[row[3].value].add_direct_team_member(row[0].value)
        sales_list[row[0].value].add_loan(int(row[6].value), row[5].value, row[4].value)

    print('Name Job Item1 Item2 Item5 OrganizationBonus')
    for i, j in sales_list.items():
        print(i, j.job, j.get_direct_service_bonus(), j.get_continuous_loan_bonus(), j.get_education_bonus(), j.get_organization_bonus())

if __name__ == '__main__':
    load_test()
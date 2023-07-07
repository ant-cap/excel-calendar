import xlsxwriter, calendar, datetime

CURRENT_DATE = datetime.date(2023, 6, 27)

CLASSES = {
    'CSE320' : {
        'assignments' : [
            [datetime.date(2023, 6, 26), 'Circuit 3 Due', True],
            [datetime.date(2023, 7, 3), 'System 4 Due', False],
            [datetime.date(2023, 7, 16), 'System 5 Due!', False]
        ],

        'color' : '#00FF00'
    },
    'CSE480' : {
        'assignments' : [
            [datetime.date(2023, 6, 29), 'Project 1 Due', False],
            [datetime.date(2023, 7, 3), 'Project 2 Due', True],
            [datetime.date(2023, 7, 13), 'Project 3 Due', True]
        ],

        'color' : '#FFFF00'
    }
}

class XlsSheet(object):
    def __init__(self):
        self.__workbook = xlsxwriter.Workbook('schedule.xlsx')
        self.__worksheet = self.__workbook.add_worksheet()

        # set borders
        border = self.__workbook.add_format({'bg_color' : 'ED7D31'})
        bleft = self.__workbook.add_format({'bg_color' : 'ED7D31', 'border_color' : '#FFFFFF', 'right' : 5 })
        bbott = self.__workbook.add_format({'bg_color' : 'ED7D31', 'border_color' : '#FFFFFF', 'top' : 5 })
        self.__worksheet.set_column(0, 0, 0.5, border)
        self.__worksheet.set_column(15, 15, 0.5, border)
        self.__worksheet.set_row(0, 4.50, border)
        self.__worksheet.set_row(43, 4.50, border)

        for i in range(1, 29):
            self.__worksheet.write(i, 0, '', bleft)
        for i in range(1, 15):
            self.__worksheet.write(29, i, '', bbott)

    def get_sheet(self):
        return self.__worksheet
    
    def get_book(self):
        return self.__workbook

    def close(self):
        self.__workbook.close()

class XlsSchedule(object):
    def __init__(self, sheet:XlsSheet, dates:list):
        self.__sheet = sheet
        self.__nodes = {}
        self.__formats = self.format1()

        x=1
        y=1
        for i in range(len(dates)):
            self.__nodes[dates[i]] = XlsNode(sheet, [x,y], dates[i], self.__formats)
            if (i+1) % 7 == 0 and i != 0:
                x += 7
                y = -1
            y += 2

    def format1(self):
        '''
        Terminal Window Style
        '''
        book = self.__sheet.get_book()
        format_dict = {}

        ##
        # Header Formats:
        #   past, current, future
        ##
        header_base1 = {
            'bold'          : True,
            'font_name'     : 'Courier New',
            'font_size'     : 12,
            'font_color'    : '#0D0D0D',
            'bg_color'      : '#808080',
            'border_color'  : '#FFFFFF',
            'top'           : 5,
            'right'         : 5
        }
        header_base2 = {
            'bold'          : True,
            'font_name'     : 'Courier New',
            'font_size'     : 9,
            'font_color'    : '#0D0D0D',
            'bg_color'      : '#808080',
            'border_color'  : '#FFFFFF',
            'right'         : 5,
            'center_across' : True
        }
        format_dict['header_past1'] = book.add_format(header_base1)
        format_dict['header_past2'] = book.add_format(header_base2)
        format_dict['header_curr1'] = book.add_format(header_base1)
        format_dict['header_curr1'].set_bg_color('#287444')
        format_dict['header_curr2'] = book.add_format(header_base2)
        format_dict['header_curr2'].set_bg_color('#287444')
        format_dict['header_futu1'] = book.add_format(header_base1)
        format_dict['header_futu1'].set_bg_color('#FFFFFF')
        format_dict['header_futu2'] = book.add_format(header_base2)
        format_dict['header_futu2'].set_bg_color('#FFFFFF')

        ##
        # Assignment formats:
        #   done, fail, due
        ##
        assignment_base = {
            'bold'          : True,
            'font_name'     : 'Courier New',
            'font_size'     : 12,
            'font_color'    : '#D9D9D9',
            'bg_color'      : '#0D0D0D',
            'border_color'  : '#FFFFFF',
            'right'         : 5
        }
        format_dict['assign_due']  = book.add_format(assignment_base)
        format_dict['assign_done'] = book.add_format(assignment_base)
        format_dict['assign_done'].set_italic()
        format_dict['assign_done'].set_font_color('#808080')
        format_dict['assign_fail'] = book.add_format(assignment_base)
        format_dict['assign_fail'].set_italic()
        format_dict['assign_fail'].set_font_color('#FA0000')

        # puts assignment_base in the dictionary for get_class_formats
        assignment_base.pop('right')
        format_dict['class_base'] = assignment_base
        format_dict['no_class'] = book.add_format(assignment_base)

        ##
        # Notes format
        ##
        format_dict['notes'] = book.add_format({
            'font_name'     : 'Courier New',
            'font_size'     : 10,
            'font_color'    : '#D9D9D9',
            'bg_color'      : '#0D0D0D',
            'border_color'  : '#FFFFFF',
            'right'         : 5,
            'valign'        : 'top'
        })

        return self.get_class_formats(format_dict)

    def get_class_formats(self, fdict:dict):
        '''
        Given the class dictionary (currently global),
        determines the format for each specific class and adds it
        to the format dictionary fdict.

        returns fdict.
        '''
        book = self.__sheet.get_book()
        class_format = fdict['class_base']

        for CLASS in CLASSES:
            color = CLASSES[CLASS]['color']
            if CLASS in fdict:
                raise Exception('class already exists')
            fdict[CLASS] = book.add_format(class_format)
            fdict[CLASS].set_font_color(color)
            fdict[CLASS].set_align('right')
        
        return fdict


    def set_assignments(self, classes:dict):
        for CLASS in classes:
            for assign in classes[CLASS]['assignments']:
                print('assignment date:', assign[0])
                if assign[0] in self.__nodes:
                    print('adding assignment')
                    self.__nodes[assign[0]].add_assignment(CLASS, assign[1], assign[2])

        for node in self.__nodes:
            self.__nodes[node].update_assignments()

class XlsNode(object):
    def __init__(self, sheet:XlsSheet, origin:list, date:datetime.date, formats):
        self.__sheet = sheet
        self.__origin = origin
        self.__date = date
        self.__formats = formats
        self.__assigns = []
        self.build_node()

    def build_node(self):
        '''
        Builds the node at the provided origin.
        This node claims this origin location, and
        formats the 7x3 block.
        '''

        def form_header():
            if CURRENT_DATE > self.__date:
                return (fdict['header_past1'], fdict['header_past2'])
            elif CURRENT_DATE < self.__date:
                return (fdict['header_futu1'], fdict['header_futu2'])
            return (fdict['header_curr1'], fdict['header_curr2'])

        x,y = self.__origin
        fdict = self.__formats
        sheet = self.__sheet.get_sheet()

        sheet.set_column(y, y, 9.30)
        sheet.set_column(y+1, y+1, 27.35)
        sheet.set_row(x, 16.50)
        sheet.set_row(x+1, 10.50)
        for i in range(1, 5):
            sheet.set_row(x+1+i, 13.50)
            sheet.write(x+1+i, y+1, '', fdict['assign_due'])
        sheet.set_row(x+6, 55.50)

        # merge day of the week + footer
        header = form_header()
        sheet.merge_range(x,y,x,y+1,self.__date.strftime('%A'), header[0])
        sheet.merge_range(x+1, y, x+1, y+1, self.__date.strftime('%d-%b'), header[1])

        #merge last row
        sheet.merge_range(x+6,y, x+6, y+1, "", fdict['notes'])
    
    def update_assignments(self):

        def form_assign(done:bool):
            if done:
                return fdict['assign_done']
            if CURRENT_DATE > self.__date and not done:
                return fdict['assign_fail']
            return fdict['assign_due']

        x,y = self.__origin
        fdict = self.__formats
        sheet = self.__sheet.get_sheet()
        assigns = self.__assigns

        for i in range(4):
            if i in range(len(assigns)):
                print('writing:', x, y)
                sheet.write(x+2+i, y, assigns[i][0], fdict[assigns[i][0]])
                sheet.write(x+2+i, y+1, assigns[i][1], form_assign(assigns[i][2]))
            else:
                sheet.write(x+2+i, y, '', fdict['no_class'])

    def add_assignment(self, course:str, name:str, done:bool):
        print('appending assignment', [course,name,done])
        self.__assigns.append([course, name, done])



def main():
    month = int(str(CURRENT_DATE)[5:7])
    year = int(str(CURRENT_DATE)[:4])
    cal = [i for i in calendar.Calendar().itermonthdates(year, month)]

    # this will only place four weeks in the schedule (a full screen)
    num_weeks = len(cal) // 7
    cd_week = (cal.index(CURRENT_DATE) // 7) + 1
    if cd_week <= 2:  # within first two weeks
        cal = cal[:28]
    elif num_weeks - cd_week <= 1:  # within last two weeks
        cal = cal[-28:]
    else:  # get current day on second row
        cal = cal[(cd_week-2)*7:((cd_week-2)*7)+28]

    xls_sheet = XlsSheet()
    schedule = XlsSchedule(xls_sheet, cal)
    schedule.set_assignments(CLASSES)

    xls_sheet.close()
    print("WOOOOOOOOW!!!")


if __name__ == "__main__":
    main()
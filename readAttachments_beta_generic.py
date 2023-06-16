from pathlib import Path
import funcoes_auxiliares_generic as func
import pandas as pd
import re

# global variables that store file and folder paths
number_patterns = ['\d\d\d\d\.\d\d\d\d', '\d\d\d\d\-\d\d\d\d',
                      '\d\d\d\d\.\d\d\d', '\d\d\d\d\-\d\d\d', 
                      '\d\d\d\d\.\d\d', '\d\d\d\d\-\d\d']

stage_patterns = ['Loren', 'Ipsum', 'Dolor']
current_path = Path.cwd()

before_general = Path(current_path).parents[5]
general = before_general / Path('Sit')
controls = general / Path('Amet/Consectetur')
obsolete = controls / Path('Adipiscing')
home_folder = general / Path('Elit/Etiam/Eget')
processed = home_folder / Path('Ligula')
unprocessed = home_folder / Path('Eu')
exceptions = home_folder / Path('Lectus')
archive = home_folder / Path('Lobortis')
parameters_file = home_folder / Path('Condimentum.txt')
appendix_wb = home_folder / Path('Aliquam.xlsx')

wb_name = 'TemplateWb'
wb = func.load_document(home_folder, wb_name, 'xlsx') # load the worksheet
ws = wb['Sheet1'] # select the correct worksheet tab
template = pd.read_excel(appendix_wb) # reads file 'appendix_wb' as a spreadsheet
descriptions = list(template['Descricao'])
template['Descricao'] = template['Descricao'].map(func.string_iterator) 
# adds a backslash to special characters in cells in the "Description" column
descriptions = template['Descricao']
count = 0
row = 1

for file in list(unprocessed.iterdir()): # iterates over the 'unprocessed' list
    if file.suffix != '.docm' and file.suffix != '.docx' and file[:1] != '20':
        file.replace(exceptions / Path(file.name))
        print('Arquivo não é do formato desejado')
        continue
        
    try:
        print(file.name)

        if file.suffix == '.docm': # checks if the file is of type 'docm'
            Nonummy = True
        else:
            Nonummy = False

        count += 1
        row += 1

        change_subject = file.name
        change_stem = file.stem
        cn_string = func.get_string_from_doc(change_subject, unprocessed)
        imp_date, imp_match = func.get_implementation_date(cn_string, Nonummy)

        if imp_match:
            imp_comment = ''
        else:
            imp_comment = imp_date
            imp_date = ''

        cn_data = ['',
                   func.get_cn_number(change_subject, number_patterns), 
                   func.get_division(change_subject, number_patterns, stage_patterns), 
                   change_stem, 
                   func.get_reference(cn_string, Nonummy), 
                   func.get_stage(change_subject, stage_patterns), 
                   'Active', '',
                   func.get_product(cn_string, Nonummy), 
                   func.get_description(cn_string, Nonummy), 
                   func.get_receipt_date(change_subject),
                   '', 
                   imp_date[:-4] + imp_date[-2:], 
                   imp_comment[:6] + imp_comment[8:], 
                   '', '', '', 'nan', '', '', '', '', '', '', '', '', '']

        if Nonummy == False:
            itemRegex = re.compile(r'(?<=Appendix\: All Selections).*')
            mo = itemRegex.search(cn_string)
            try:
                cn_string = mo.group()
                doc_start = 0
            except:
                file.replace(exceptions / Path(change_subject))
                print('Erro no appendix')

            try:
                for i, desc in enumerate(descriptions):
    
                    if i == 15:
                        itemRegex = re.compile(fr'({desc})(.*)')
                        mo = itemRegex.search(cn_string[doc_start:])
                    elif i == 0:
                        itemRegex = re.compile(fr'({desc})(.*?)({descriptions[i + 1]})')
                        mo = itemRegex.search(cn_string)
                        doc_start = len(mo.group(1)) + len(mo.group(2))
                    else:
                        itemRegex = re.compile(fr'({desc})(.*?)({descriptions[i + 1]})')
                        mo = itemRegex.search(cn_string[doc_start:doc_start + 2000])
                        doc_start += len(mo.group(1)) + len(mo.group(2))
    
                    descricao = mo.group(1).strip()
                    answer = func.find_x(mo.group(2).strip())
                    cn_data.append(answer)

            except:
                file.replace(exceptions / Path(change_subject))
                print('Erro nas perguntas')

 
        for column_number, column in enumerate(cn_data):
                ws.cell(row, column=column_number+1, value=column)
        
        file.replace(processed / Path(change_subject))            

    except:
        file.replace(exceptions / Path(change_subject))
        print('Erro não identificado, por favor insira esses dados manualmente')
        pass

wb.save(home_folder / Path('Auctor.xlsx')) # update spreadsheet

end = str(input('Bot executado com sucesso!'))

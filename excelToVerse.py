from openpyxl import load_workbook
from stringcase import pascalcase, camelcase

class Parser:
    def __init__(self, load_ws, load_meta, sheetName):
        self.load_ws = load_ws
        self.load_meta = load_meta
        self.dataList = []
        
        sheetName = sheetName[1:-1]
        self.sheetName = str(sheetName)

        self.parsePath = load_meta.cell(1, 2).value
        self.closeConsole = load_meta.cell(2, 2).value
        
        self.all_values = []
        self.field_names = []
        self.field_types = []
        self.item_field_dict = []
        passFieldName = False
        passFieldType = False
        
        dataList = self.dataList
        all_values = self.all_values
        field_names = self.field_names
        field_types = self.field_types
        item_field_dict = self.item_field_dict
        
        for row in load_ws.rows:
            row_value = []
            for cell in row:
                row_value.append(cell.value)
                
            if(passFieldName and passFieldType):
                all_values.append(row_value)
            
            if(passFieldName and not passFieldType):
                passFieldType = True
                field_types = row_value
                
            if(not passFieldName):
                passFieldName = True
                field_names = row_value
        
        for i in range(len(field_names)):
            name = field_names[i]
            type = field_types[i]
            item_field_dict.append({"name":name, "type":type})
        
        self.fieldKey = item_field_dict[0]['name']
        self.mapKey = item_field_dict[1]['name']
        
        for item in all_values:
            item_dict = {}
            
            for i in range(len(field_names)):
                field = field_names[i]
                item_dict[field] = item[i]
            dataList.append(item_dict)


    def get_type_initValue(self, field_type):
        if(field_type == "string"):
            return "\"\""
        if(field_type == "int"):
            return "0"
        if(field_type == "float"):
            return "0.0"
        if(field_type == "logic"):
            return "false"  

        print("get_type_initValue error")
        return "ERROR!!!"   

    def get_wrapped_value(self, value):
        if(isinstance(value, int)):
            return value
        if(isinstance(value, float)):
            return value
        if(isinstance(value, str)):
            return "\"{value}\"".format(value = value)
        if(isinstance(value, bool)):
            return str(value).lower()

        return value    
    
    def get_value_type(self, value):
        if(isinstance(value, int)):
            return "int"
        if(isinstance(value, float)):
            return "float"
        if(isinstance(value, str)):
            return "string"
        if(isinstance(value, bool)):
            raise Exception("logic(bool)은 키 값으로 사용할 수 없습니다.")
        

    def get_Item_template(self, fieldList):
        template = """custom_{name}_item := class:
""".format(name=self.sheetName)

        fieldTemplate = "    var {name}<public>: {type} = {initValue}\n"

        for field in fieldList:
            if(field['name'] == self.fieldKey):
                continue
            if(field['name'] == self.mapKey):
                continue
            template += fieldTemplate.format(name=pascalcase(field['name']), type=field['type'], initValue = self.get_type_initValue(field['type']))  

        return template 

    def get_make_item_template(self, fieldList):
        template = "make_custom_{name}_item<constructor>(".format(name=self.sheetName)

        for i in range(len(fieldList)):
            field = fieldList[i]
            
            if(field['name'] == self.fieldKey):
                continue
            if(field['name'] == self.mapKey):
                continue
            
            template += "Arg{number}: {type}".format(number=i, type=field['type'])
            if(i < len(fieldList) - 1):
                template += ", "
        template += ") := custom_{name}_item:\n".format(name=self.sheetName)

        fieldTemplate = "    {name} := Arg{number}\n"   

        for i in range(len(fieldList)):
            field = fieldList[i]
            
            if(field['name'] == self.fieldKey):
                continue
            if(field['name'] == self.mapKey):
                continue
            
            template += fieldTemplate.format(name=pascalcase(field['name']), number=i) 

        return template 

    def get_data_template(self, fieldKey, dataList):
        template = """custom_{name}_data := class:
""".format(name=self.sheetName)
        mapTemplate = """    var Table<public>: [{keyType}]custom_{type}_item = map{{}}\n
    TableCreate<public>():void=
""".format(type = self.sheetName, keyType=self.item_field_dict[1]["type"])

        for data in dataList:
            print(data)
            template += "    var {weapon}<public>:custom_{name}_item = ".format(weapon=pascalcase(data[fieldKey]), name=self.sheetName)
            template += "custom_{name}_item{{}}\n".format(name=self.sheetName)
            mapTemplate += "        if(set Table[{key}] = {item}) {{}}\n".format(key = self.get_wrapped_value(data[self.mapKey]), item = pascalcase(data[self.fieldKey]))

        return template + "\n" + mapTemplate

    def get_data_contstructor(self, fieldKey, dataList):
        template = """make_custom_{name}_data<constructor>() := custom_{name}_data:
""".format(name=self.sheetName)
        for data in dataList:
            template += "    {name} := make_custom_{classname}_item(".format(name=pascalcase(data[fieldKey]), classname=self.sheetName)

            vs = [*data.values()]
            for i in range(len(vs)):
                if(i == 0 or i == 1):
                    continue
                
                value = vs[i]
                template += "{value}".format(value = self.get_wrapped_value(value))
                if(i < len(vs)-1):
                    template += ", "
            template += ")\n"
        return template
    
    def Parse(self):
        parseItemTemplate = self.get_Item_template(self.item_field_dict)
        parseItemConstructorTemplate = self.get_make_item_template(self.item_field_dict)
        parseDataTemplate = self.get_data_template(self.fieldKey, self.dataList)
        parseDataConstructorTemplate = self.get_data_contstructor(self.fieldKey, self.dataList)

        print(parseItemTemplate)
        print(parseItemConstructorTemplate)
        print(parseDataTemplate)
        print(parseDataConstructorTemplate)
        s = ""
        s += parseItemTemplate + "\n"
        s += parseItemConstructorTemplate + "\n"
        s += parseDataTemplate + "\n"
        s += parseDataConstructorTemplate + "\n"
        
        return s
        



load_wb  = load_workbook("dataTable.xlsm", data_only = True)
load_meta = load_wb["_meta_"]

parsePath = load_meta.cell(1, 2).value
closeConsole = load_meta.cell(2, 2).value

s = ""

for sheet in load_wb.sheetnames:
    load_ws = None
    
    if(len(sheet) == 0):
        continue
    if(not (sheet[0] == '_' and sheet[-1] == '_')):
        continue
    if(sheet == "_meta_"):
        continue
    
    try:
        load_ws = load_wb[sheet]
    except:
        print("sheet load failed")
        pass
    
    try:
        print("-----------try to parse {name}-----------".format(name=sheet))
        print()
        p = Parser(load_ws, load_meta, sheet)
        s += "\n#=============================================={name}==============================================\n".format(name = sheet)
        s += p.Parse()
        s += "\n#==============================================\n#==============================================\n\n"
    except Exception as e:
        print(sheet + " sheet parse faield")
        print(e)
        
try:
    with open(parsePath+"\data.verse", 'w') as writer:
       writer.write(s + "\n")
except:
    print("file writing error")
        
if(not closeConsole):
    input("exit ?")

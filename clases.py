class Element_Excel: # representa un elmento en l archivo excel

    def __init__(self,x,y) -> None:
        self.position = {"x": x, "y": y}
    
    def get_position(self):
        return self.position["x"], self.position["y"]

    def set_position(self, x, y):
        self.position = {"x": x, "y": y}

class Element_txt(Element_Excel): # represnta un elmnto de tipo texto de el archivo excel 

    def __init__(self, x, y, val, size, font, col) -> None:
        super().__init__(x, y)
        self.text = val
        self.size = size
        self.font = font
        self.color = col


class Elemnt_img(Element_Excel): # represnta un elmnto de tipo img de el archivo excel 

    def __init__(self, x, y, rute, size) -> None:
        super().__init__(x, y)
        self.rute = rute
        self.size = size

class Element_artic(Element_Excel):
    cod_size = 8
    desc_size = 75
    unit_size = 5



    def __init__(self, x, y, code, desc, unit, price) -> None:
        super().__init__(x, y)
        self.code = Element_txt(x,y,code,14,"arial", "green")
        self.description = Element_txt(x,y+self.cod_size,desc,12,"arial", "black")
        self.unit = Element_txt(x,y+self.desc_size,unit,14,"arial", "black")
        self.price = Element_txt(x,y+self.unit_size,price,14,"arial", "red")

    def get_code(self):
        return self.code.text
    
    def set_code(self, code):
        self.code.text = code

    def get_description(self):
        return self.description.text
    
    def set_description(self, desc):
        self.description.text = desc
    
    def get_unit(self):
        return self.unit.text
    
    def set_unit(self, unit):
        self.unit.text = unit
    
    def get_price(self):
        return self.price.text
    
    def set_price(self, price):
        self.price.text = price


class Arch_excel:

    def __init__(self, f_mod, rute, elements = []) -> None:
        self.f_mod = f_mod
        self.rute = rute
        self.elements = elements

    def append_element(self, element):
        self.elements.append(element)

    def __iter__(self):
        for i in range(0, len(self.elements)):
            yield self.elements[i]
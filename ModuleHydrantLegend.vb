Module ModuleHydrantLegend

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Determine the legend code for 1 hydrant, 
    '''     based on all relevant attribute values.
    ''' </summary>
    ''' <param name="StatusCode">string</param>
    ''' <param name="hydrantTypeCode">string</param>
    ''' <param name="LiggingCode">string</param>
    ''' <param name="Diameter">integer</param>
    ''' <returns>integer</returns>
    ''' <remarks>
    '''     Decision algorithm based on rules defined in configuration file.
    ''' </remarks>
    ''' <history>
    ''' 	[ex00764]	8/08/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function HydrantLegendCode( _
        ByVal StatusCode As String, _
        ByVal hydrantTypeCode As String, _
        ByVal LiggingCode As String, _
        ByVal Diameter As Integer _
        ) As Integer

        Try

            Dim Code As String = "" 'the final value to return
            Dim TmpStr As String 'temporary string used for holding ini-file values
            Dim Rules As String() 'array of rule keys from ini-file
            Dim RuleDef As String() 'array of (1 or more) values for one rule
            Dim NumberOfRules As Integer
            Dim i As Integer 'loop index

            'Get the list of decision rules from the Config.ini file.
            TmpStr = INIRead(g_FilePath_Config, "HydrantenLegendRules") ' get all keys in section
            TmpStr = TmpStr.Replace(ControlChars.NullChar, "|"c) ' change embedded NULLs to pipe chars
            Rules = TmpStr.Split("|")
            NumberOfRules = Rules.Length

            'Loop through this list, one by one, until a matching rule is found.
            For i = 0 To NumberOfRules - 1

                'Retrieve the definition of the current rule.
                TmpStr = INIRead(g_FilePath_Config, "HydrantenLegendRules", Rules(i))
                RuleDef = TmpStr.Split(";")

                'Check the Status condition of current rule.
                If StatusCode = RuleDef(0) Then
                    'Check if there is a second condition.
                    If RuleDef.Length > 1 Then

                        'Check the HydrantType condition of current rule.
                        If hydrantTypeCode = RuleDef(1) Then
                            'Check if there is a third condition.
                            If RuleDef.Length > 2 Then

                                'Check the Ligging condition of current rule.
                                If LiggingCode = RuleDef(2) Then
                                    'Rule does match. Continue with legend code of current rule.
                                    Code = Rules(i)
                                    Exit For
                                End If

                            Else 'There is no third condition.
                                'Rule does match. Continue with legend code of current rule.
                                Code = Rules(i)
                                Exit For
                            End If
                        End If

                    Else 'There is no second condition.
                        'Rule does match. Continue with legend code of current rule.
                        Code = Rules(i)
                        Exit For
                    End If
                End If

            Next

            'Does the legend code require any processing ?
            'Return the resulting value.
            Try
                Select Case Code
                    Case "0k+d"
                        HydrantLegendCode = CInt(Diameter)
                    Case "1k+d"
                        HydrantLegendCode = 1000 + CInt(Diameter)
                    Case "2k+d"
                        HydrantLegendCode = 2000 + CInt(Diameter)
                    Case "3k+d"
                        HydrantLegendCode = 3000 + CInt(Diameter)
                    Case "4k+d"
                        HydrantLegendCode = 4000 + CInt(Diameter)
                    Case "5k+d"
                        HydrantLegendCode = 5000 + CInt(Diameter)
                    Case "6k+d"
                        HydrantLegendCode = 6000 + CInt(Diameter)
                    Case "7k+d"
                        HydrantLegendCode = 7000 + CInt(Diameter)
                    Case "8k+d"
                        HydrantLegendCode = 8000 + CInt(Diameter)
                    Case "9k+d"
                        HydrantLegendCode = 9000 + CInt(Diameter)
                    Case Else
                        HydrantLegendCode = CInt(Code)
                End Select
            Catch
                'Use default legend code value in case no matching rule 
                'could be found, or in case of error.
                HydrantLegendCode = 0
            End Try

        Catch ex As Exception
            Throw ex
        End Try

    End Function

End Module

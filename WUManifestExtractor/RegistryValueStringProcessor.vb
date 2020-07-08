Module RegistryValueStringProcessor
    Public Function SplitContinuousBinaryStirng(ByVal BinaryString As String, Optional AddZeroAtLast As Boolean = False, Optional SeparateCharacter As String = ",") As String
        Dim SplittedBinaryString As String = ""
        If BinaryString.Length = 0 Then
            SplittedBinaryString = ""
            Return SplittedBinaryString
        End If
        If BinaryString.Length Mod 2 = 1 Then
            If AddZeroAtLast Then
                BinaryString = BinaryString & "0"
            Else
                BinaryString = "0" & BinaryString
            End If
        End If
        For iIndex = 0 To BinaryString.Length - 2 Step 2
            SplittedBinaryString = SplittedBinaryString & BinaryString.Substring(iIndex, 2)
            If iIndex <> BinaryString.Length - 2 Then
                SplittedBinaryString = SplittedBinaryString & SeparateCharacter
            End If
        Next
        Return SplittedBinaryString
    End Function
    Public Function StringToRegistryStringBinaryValue(ByVal StringValue As String) As String
        Dim RegSzVal As String = ""
        Dim CharTemp As String
        If StringValue.Length = 0 Then
            RegSzVal = ""
        Else
            For iIndex = 0 To StringValue.Length - 1
                CharTemp = Convert.ToString(AscW(StringValue(iIndex)), 16)
                If CharTemp.Length = 2 Then
                    RegSzVal = RegSzVal & CharTemp & ",00,"
                Else
                    RegSzVal = RegSzVal & CharTemp.Substring(2, 2) & "," & CharTemp.Substring(0, 2) & ","
                End If
            Next
            RegSzVal = RegSzVal.Substring(0, RegSzVal.Length - 1)
        End If
        Return RegSzVal
    End Function
    Public Function ManifestRegistryExpendableStringValueToRegFileExpendableStringValue(ByVal ManifestRegistryExpendableStringValue As String) As String
        Dim RegExpandSzVal As String = "hex(2):"
        Dim CharTemp As String
        If ManifestRegistryExpendableStringValue.Length = 0 Then
            RegExpandSzVal = RegExpandSzVal & "00,00"
        Else
            For iIndex = 0 To ManifestRegistryExpendableStringValue.Length - 1
                CharTemp = Convert.ToString(AscW(ManifestRegistryExpendableStringValue(iIndex)), 16)
                If CharTemp.Length = 2 Then
                    RegExpandSzVal = RegExpandSzVal & CharTemp & ",00,"
                Else
                    RegExpandSzVal = RegExpandSzVal & CharTemp.Substring(2, 2) & "," & CharTemp.Substring(0, 2) & ","
                End If
            Next
            RegExpandSzVal = RegExpandSzVal & "00"
        End If
        Return RegExpandSzVal
    End Function
    Public Function ManifestRegistryMultiStringValueToRegFileMultiStringValue(ByVal ManifestRegistryMultiStringValue As String)
        Dim RegMultiSzVal As String = "hex(7):"
        Dim IsSingleLineEnded As Boolean = True
        If ManifestRegistryMultiStringValue.Length = 0 Then
            RegMultiSzVal = RegMultiSzVal & "00,00"
        Else
            ManifestRegistryMultiStringValue = ManifestRegistryMultiStringValue.Replace("""""", vbLf)
            Dim ArrayRegSz() As String = ManifestRegistryMultiStringValue.Split(vbLf)
            If ArrayRegSz(0).StartsWith("""") Then
                If ArrayRegSz(0).Length <= 1 Then
                    ArrayRegSz(0) = ""
                Else
                    ArrayRegSz(0) = ArrayRegSz(0).Substring(1)
                End If
            End If
            If ArrayRegSz(ArrayRegSz.Length - 1).EndsWith("""") Then
                If ArrayRegSz(ArrayRegSz.Length - 1).Length <= 1 Then
                    ArrayRegSz(ArrayRegSz.Length - 1) = ""
                Else
                    ArrayRegSz(ArrayRegSz.Length - 1) = ArrayRegSz(ArrayRegSz.Length - 1).Substring(0, ArrayRegSz(ArrayRegSz.Length - 1).Length - 1)
                End If
            End If
            For iStrIndex = 0 To ArrayRegSz.Length - 1
                If ArrayRegSz(iStrIndex).Length = 0 Then
                    RegMultiSzVal = RegMultiSzVal & StringToRegistryStringBinaryValue(ArrayRegSz(iStrIndex)) & "00,00,"
                Else
                    RegMultiSzVal = RegMultiSzVal & StringToRegistryStringBinaryValue(ArrayRegSz(iStrIndex)) & ",00,00,"
                End If
            Next
            RegMultiSzVal = RegMultiSzVal & "00,00"
        End If
        Return RegMultiSzVal
    End Function
    Public Function ManifestRegistryStringValueCheck(ByVal ManifestStringValue As String) As String
        Dim RegSzVal As String = ManifestStringValue.Replace("\""", """")
        RegSzVal = RegSzVal.Replace("\\", "\")
        RegSzVal = RegSzVal.Replace("\", "\\")
        RegSzVal = RegSzVal.Replace("""", "\""")
        Return RegSzVal
    End Function
    Public Function ManifestRegistryStringValueToRegFileStringValue(ByVal ManifestStringValue As String) As String
        Dim RegSzVal As String = ManifestStringValue
        RegSzVal = RegSzVal.Replace("\", "\\")
        RegSzVal = RegSzVal.Replace("""", "\""")
        Return RegSzVal
    End Function
    Public Function ManifestRegistryBinaryValueToRegFileBinaryValue(ByVal ManifestRegistryBinaryValue As String) As String
        Dim RegBinVal As String = "hex:"
        If ManifestRegistryBinaryValue.Length = 0 Then
            RegBinVal = "hex:"
        Else
            If ManifestRegistryBinaryValue.Length Mod 2 = 1 Then
                ManifestRegistryBinaryValue = "0" & ManifestRegistryBinaryValue
            End If
            Dim iIndex As Integer = 0
            For iIndex = 0 To ManifestRegistryBinaryValue.Length - 2 Step 2
                RegBinVal = RegBinVal & ManifestRegistryBinaryValue.Substring(iIndex, 2)
                If iIndex <> ManifestRegistryBinaryValue.Length - 2 Then
                    RegBinVal = RegBinVal & ","
                End If
            Next
        End If
        Return RegBinVal
    End Function
End Module

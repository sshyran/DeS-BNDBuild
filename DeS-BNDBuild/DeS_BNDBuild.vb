Imports System.IO
Imports System.IO.Compression
Imports System.Numerics
Imports System.Threading
Imports System.Runtime.Serialization.Formatters.Binary
'Imports System.Security.Cryptography.RijndaelManaged
Imports System.Security.Cryptography



Public Class Des_BNDBuild
    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String
    Public Shared extractPath As String

    'Stuff for (de)swizzling
    Public Shared writeOffset As UInteger
    Public Shared skipOffset As UInteger
    Public Shared inBytes As Byte()
    Public Shared outBytes As Byte()
    Public Shared ddsWidth As UShort
    Public Shared ddsBlockArray As Boolean()

    Public Shared bigEndian As Boolean = True

    Private trdWorker As Thread

    Dim outputLock As New Object
    Dim workLock As New Object

    Public Shared work As Boolean = False
    Public Shared outputList As New List(Of String)

    Public Structure pathHash
        Dim hash As UInteger
        Dim bdtoffset As ULong
        Dim filesize As UInteger
        Dim idx As UInteger
    End Structure

    Public Structure hashGroup
        Dim length As UInteger
        Dim idx As UInteger
    End Structure

    Public Structure saltedHash
        Dim idx As UInteger
        Dim shaHash() As Byte
        Dim rangeCount As UInteger
        Dim ranges As List(Of bhd5Range)
    End Structure

    Public Structure bhd5Range
        Dim startOffset As Long
        Dim endOffset As Long
    End Structure

    Public Structure DDS_HEADER
        Dim dwSize As UInteger
        Dim dwFlags As UInteger
        Dim dwHeight As UInteger
        Dim dwWidth As UInteger
        Dim dwPitchOrLinearSize As UInteger
        Dim dwDepth As UInteger
        Dim dwMipMapCount As UInteger
        Dim dwReserved1() As UInteger
        Dim ddspf As DDS_PIXELFORMAT
        Dim dwCaps As UInteger
        Dim dwCaps2 As UInteger
        Dim dwCaps3 As UInteger
        Dim dwCaps4 As UInteger
        Dim dwReserved2 As UInteger
    End Structure

    Public Structure DDS_PIXELFORMAT
        Dim dwSize As UInteger
        Dim dwFlags As UInteger
        Dim dwFourCC As UInteger
        Dim dwRGBBitCount As UInteger
        Dim dwRBitMask As UInteger
        Dim dwGBitMask As UInteger
        Dim dwBBitMask As UInteger
        Dim dwABitMask As UInteger
    End Structure

    Public Structure DDS_HEADER_DXT10
        Dim dxgiFormat As UInteger
        Dim resourceDimension As UInteger
        Dim miscFlag As UInteger
        Dim arraySize As UInteger
        Dim miscFlags2 As UInteger
    End Structure

    Public Enum GAME_PLATFORM
        PS3
        PS4
    End Enum

    Private WithEvents updateUITimer As New System.Windows.Forms.Timer()

    Dim ShiftJISEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

    Dim UnicodeEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("UTF-16")

    Public Sub output(txt As String)
        SyncLock outputLock
            outputList.Add(txt)
        End SyncLock
    End Sub

    Private Function EncodeFileName(ByVal filename As String) As Byte()
        Return ShiftJISEncoding.GetBytes(filename)
    End Function

    Private Function EncodeFileNameBND4(ByVal filename As String) As Byte()
        Return UnicodeEncoding.GetBytes(filename)
    End Function

    Private Sub EncodeFileName(ByVal filename As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = ShiftJISEncoding.GetBytes(filename)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub

    Private Sub EncodeFileNameBND4(ByVal filename As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = UnicodeEncoding.GetBytes(filename)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub

    Private Function DecodeFileName(ByVal loc As UInteger) As String
        Dim b As New System.Collections.Generic.List(Of Byte)
        Dim cont As Boolean = True

        While cont
            If bytes(loc) > 0 Then
                b.Add(bytes(loc))
                loc += 1
            Else
                cont = False
            End If
        End While

        Return ShiftJISEncoding.GetString(b.ToArray())
    End Function

    Private Function DecodeFileNameBND4(ByVal loc As UInteger) As String
        Dim b As New System.Collections.Generic.List(Of Byte)
        Dim cont As Boolean = True

        While True
            If bytes(loc) > 0 Or bytes(loc + 1) > 0 Then
                b.Add(bytes(loc))
                b.Add(bytes(loc + 1))
                loc += 2
            Else
                Exit While
            End If
        End While

        Return UnicodeEncoding.GetString(b.ToArray())
    End Function

    Private Function StrFromBytes(ByVal loc As UInteger) As String
        Dim Str As String = ""
        Dim cont As Boolean = True

        While cont
            If bytes(loc) > 0 Then
                Str = Str + Convert.ToChar(bytes(loc))
                loc += 1
            Else
                cont = False
            End If
        End While

        Return Str
    End Function

    Private Function StrFromNumBytes(ByVal loc As UInteger, ByRef num As UInteger) As String
        Dim Str As String = ""

        For i As UInteger = 0 To num - 1
            Str = Str + Convert.ToChar(bytes(loc))
            loc += 1
        Next

        Return Str
    End Function

    Private Function UInt16FromBytes(ByVal loc As UInteger) As UShort
        Dim tmpUint As UInteger = 0
        Dim bArr(1) As Byte

        Array.Copy(bytes, loc, bArr, 0, 2)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt16(bArr, 0)

        Return tmpUint
    End Function

    Private Function UIntFromBytes(ByVal loc As UInteger) As UInteger
        Dim tmpUint As UInteger = 0
        Dim bArr(3) As Byte

        Array.Copy(bytes, loc, bArr, 0, 4)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt32(bArr, 0)

        Return tmpUint
    End Function

    Private Function UInt64FromBytes(ByVal loc As UInteger) As ULong
        Dim tmpUint As ULong = 0
        Dim bArr(7) As Byte

        Array.Copy(bytes, loc, bArr, 0, 8)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt64(bArr, 0)

        Return tmpUint
    End Function

    Private Sub StrToBytes(ByVal str As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = System.Text.Encoding.ASCII.GetBytes(str)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub
    Private Function StrToBytes(ByVal str As String) As Byte()
        'Return bytes of string, do not insert
        Return System.Text.Encoding.ASCII.GetBytes(str)
    End Function

    Private Sub InsBytes(ByVal bytes2() As Byte, ByVal loc As Long)
        Array.Copy(bytes2, 0, bytes, loc, bytes2.Length)
    End Sub

    Private Sub UInt16ToBytes(ByVal val As UInt16, loc As UInteger)

        Dim bArr(1) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, 2)
    End Sub

    Private Sub UIntToBytes(ByVal val As UInteger, loc As UInteger)

        Dim bArr(3) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, 4)
    End Sub

    Private Sub UInt64ToBytes(ByVal val As ULong, loc As UInteger)

        Dim bArr(7) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, 8)
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim openDlg As New OpenFileDialog()

        openDlg.Filter = "DS DCX/BND File|*BND;*MOWB;*DCX;*TPF;*BHD5;*BHD;*SL2;*BDLE;*BDLEDEBUG;*ENTRYFILELIST"
        openDlg.Multiselect = True
        openDlg.Title = "Open your BND files"

        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtBNDfile.Lines = openDlg.FileNames
        End If


    End Sub


    Private Function HashFileName(filename As String) As UInteger

        REM This code copied from https://github.com/Burton-Radons/Alexandria

        If filename Is Nothing Then
            Return 0
        End If

        Dim hash As UInteger = 0

        For Each ch As Char In filename
            hash = hash * &H25 + Asc(Char.ToLowerInvariant(ch))
        Next

        Return hash
    End Function
    Private Function DecryptSoundFile(ByRef soundBytes As Byte()) As Byte()
        Dim key As Byte() = System.Text.Encoding.ASCII.GetBytes("G0KTrWjS9syqF7vVD6RaVXlFD91gMgkC")
        Dim iv(15) As Byte ' = soundBytes.Take(16).ToArray()

        Dim encBytes As Byte() = soundBytes

        Dim ms As New MemoryStream()
        Dim aes As New AesManaged() With {
            .Mode = CipherMode.CBC,
            .Padding = PaddingMode.Zeros
        }
        Dim cs As New CryptoStream(ms, aes.CreateDecryptor(key, iv), CryptoStreamMode.Write)

        cs.Write(encBytes, 0, encBytes.Length)

        cs.Dispose()
        Return ms.ToArray()
    End Function
    Private Function DecryptRegulationFile(ByRef regBytes As Byte()) As Byte()
        Dim key As Byte() = System.Text.Encoding.ASCII.GetBytes("ds3#jn/8_7(rsY9pg55GFN7VFL#+3n/)")
        Dim iv As Byte() = regBytes.Take(16).ToArray()

        Dim encBytes As Byte() = regBytes.Skip(16).ToArray()

        Dim ms As New MemoryStream()
        Dim aes As New AesManaged() With {
            .Mode = CipherMode.CBC,
            .Padding = PaddingMode.Zeros
        }
        Dim cs As New CryptoStream(ms, aes.CreateDecryptor(key, iv), CryptoStreamMode.Write)

        cs.Write(encBytes, 0, encBytes.Length)

        cs.Dispose()
        Return ms.ToArray()
    End Function
    Private Function EncryptRegulationFile(ByRef regBytes As Byte()) As Byte()
        Dim key As Byte() = System.Text.Encoding.ASCII.GetBytes("ds3#jn/8_7(rsY9pg55GFN7VFL#+3n/)")

        Dim ms As New MemoryStream()
        Dim aes As New AesManaged() With {
            .Mode = CipherMode.CBC,
            .Padding = PaddingMode.Zeros
        }
        Dim iv(15) As Byte

        Dim cs As New CryptoStream(ms, aes.CreateEncryptor(key, iv), CryptoStreamMode.Write)

        cs.Write(regBytes, 0, regBytes.Length)

        cs.FlushFinalBlock()

        Dim encBytes(iv.Length + ms.Length - 1) As Byte

        Array.Copy(ms.ToArray, 0, encBytes, 16, ms.Length)

        Return encBytes
    End Function
    Private Sub WriteBytes(ByRef fs As FileStream, ByVal byt() As Byte)
        'Write to stream at present location
        For i = 0 To byt.Length - 1
            fs.WriteByte(byt(i))
        Next
    End Sub


    Private Sub BtnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        trdWorker = New Thread(AddressOf Extract)
        trdWorker.IsBackground = True
        trdWorker.Start()
    End Sub
    Private Sub Extract()
        SyncLock workLock
            work = True
        End SyncLock

        'Dim chars As String = "abcdefghijklmnopqrstuvwxyz_-."
        'Dim chars2 As String = "abcdefghijklmnopqrstuvwxyz_-1234567890."

        'For k = 0 To chars.Length - 1


        '    For l = 0 To chars.Length - 1

        '        For m = 0 To chars.Length - 1

        '            For n = 0 To chars.Length - 1

        '                For o = 0 To chars.Length - 1
        '                    'If HashFileName("/font/" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & "_map/font.gfx") = &H8486AB34 Then
        '                    '    MsgBox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o))
        '                    'End If
        '                    HashFileName(chars(k) & chars(l) & chars(m) & chars(n) & chars(o))
        '                    'If HashFileName("/font/" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & "_std/font.gfx") = &H8486AB34 Then
        '                    '    MsgBox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o))
        '                    'End If
        '                    'If HashFileName("/font/" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & "_texteffect/font.gfx") = &H8486AB34 Then
        '                    '    MsgBox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o))
        '                    'End If
        '                    For p = 0 To chars.Length - 1
        '                        'If HashFileName("/font/rusru_" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p) & "/font.gfx") = &H8486AB34 Then
        '                        '    MsgBox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p))
        '                        'End If
        '                        For q = 0 To chars.Length - 1
        '                            If HashFileName(chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p) & chars(q)) = &H4528C031 Then
        '                                MsgBox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p) & chars(q))
        '                            End If
        '                        Next

        '                        '                    'if hashfilename("/parts/" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p) & "/bd_a_underwear.dds") = &h8b4c5902 then
        '                        '                    if hashfilename("/other/" & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p) & ".dcx") = &h8b4c5902 then
        '                        '                        msgbox("found it: " & chars(k) & chars(l) & chars(m) & chars(n) & chars(o) & chars(p))
        '                        '                    end if
        '                    Next
        '                Next
        '            Next
        '        Next
        '    Next
        '    output("once" & Environment.NewLine)
        'Next

        'For k = 0 To chars2.Length - 1
        '    For l = 0 To chars2.Length - 1
        '        For m = 0 To chars2.Length - 1
        '            For n = 0 To chars2.Length - 1
        '                For o = 0 To chars2.Length - 1
        '                    'If HashFileName("/map/m17_00_00_00/m17_00_00_00_301010.mapbnd.dcx" & chars2(k) & chars2(l) & chars2(m) & chars2(n) & chars2(o)) = &HA28B84C8 Then
        '                    '    MsgBox("found it: " & chars2(k)) '& chars2(l) & chars2(m) & chars2(n) & chars2(o))
        '                    'End If
        '                    If HashFileName("/map/m17_00_00_00/m17_00_00_00_301010.mapbnd.dcx" & chars2(k) & chars2(l) & chars2(m) & chars2(n) & chars2(o)) = &HA28B84C8 Then
        '                        MsgBox("found it: " & chars2(k)) '& chars2(l) & chars2(m) & chars2(n) & chars2(o))
        '                    End If
        '                Next
        '            Next
        '        Next
        '    Next
        'Next
        'output(TimeOfDay & " - hash." & HashFileName(txtBNDfile.Lines(0)) & Environment.NewLine)
        ''output("done" & Environment.NewLine)
        'SyncLock workLock
        '    work = False
        'End SyncLock

        'Return

        Try

            For Each bndfile In txtBNDfile.Lines
                'TODO:  Confirm endian correctness for all DeS/DaS PC/PS3 formats
                'TODO:  Bitch about the massive job that is the above

                'TODO:  Do it anyway.

                'TODO:  In the endian checks, look into why you check if it equals 0
                '       Since that can't matter, since a non-zero value will still
                '       be non-zero in either endian.

                '       Seriously, what the hell were you thinking?


                bigEndian = True
                Dim DCX As Boolean = False
                Dim IsRegulation As Boolean = False

                Dim currFileName As String = ""
                Dim currFilePath As String = ""
                Dim fileList As String = ""

                Dim BinderID As String = ""
                Dim namesEndLoc As UInteger = 0
                Dim flags As UInteger = 0
                Dim numFiles As UInteger = 0

                filepath = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, "\"))
                filename = Microsoft.VisualBasic.Right(bndfile, bndfile.Length - filepath.Length)
                Try
                    bytes = File.ReadAllBytes(filepath & filename)
                Catch ex As Exception
                    MsgBox(ex.Message, MessageBoxIcon.Error)
                    SyncLock workLock
                        work = False
                    End SyncLock
                    Return
                End Try

                output(TimeOfDay & " - Beginning extraction." & Environment.NewLine)


                If Microsoft.VisualBasic.Right(filename, 3) = "bhd" Then
                    Dim firstBytes As UInteger = UIntFromBytes(&H0)

                    If archiveDict.ContainsKey(firstBytes) Then
                        Dim archiveName = archiveDict(firstBytes)

                        If archiveName = "Data0" Then
                            Try
                                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                                bytes = File.ReadAllBytes(filepath & filename)
                            Catch ex As Exception
                                MsgBox(ex.Message, MessageBoxIcon.Error)
                                SyncLock workLock
                                    work = False
                                End SyncLock
                                Return
                            End Try

                            output(TimeOfDay & " - Beginning decryption of regulation file." & Environment.NewLine)

                            bytes = DecryptRegulationFile(bytes)

                            output(TimeOfDay & " - Finished decryption of regulation file." & Environment.NewLine)

                            IsRegulation = True
                        Else
                            If Not File.Exists(filepath & ".enc.bak") Then
                                File.WriteAllBytes(filepath & filename & ".enc.bak", bytes)
                                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                                output(TimeOfDay & " - " & filename & ".enc.bak created." & Environment.NewLine)
                            Else
                                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                                output(TimeOfDay & " - " & filename & ".enc.bak already exists." & Environment.NewLine)
                            End If
                            output(TimeOfDay & " - Beginning decryption." & Environment.NewLine)
                            Dim decStream
                            If SekiroRadio.Checked Then
                                decStream = New IO.FileStream(filepath & filename, IO.FileMode.Create)
                            Else
                                decStream = New IO.FileStream(filepath & archiveName & ".bhd", IO.FileMode.Create)
                            End If

                            Dim idxSheet As ULong = 0
                            Dim diff As UInteger = 0
                            Dim countSheet As UInteger = 256
                            Dim exp As New BigInteger(BitConverter.ToUInt32(expDict(archiveName), 0))
                            Dim modulus As New BigInteger(modDict(archiveName))

                            While idxSheet < bytes.Length
                                diff = bytes.Length - idxSheet
                                If diff < 256 Then
                                    countSheet = diff
                                End If
                                Dim tempBlock As Byte() = (bytes.Skip(idxSheet).Take(countSheet)).ToArray()
                                Array.Reverse(tempBlock)
                                ReDim Preserve tempBlock(tempBlock.Length)

                                Dim processBlock As New BigInteger(tempBlock)

                                processBlock = BigInteger.ModPow(processBlock, exp, modulus)

                                Dim processBlockBytes As Byte() = processBlock.ToByteArray().Reverse().ToArray()

                                Dim padding = (countSheet - 1) - processBlockBytes.Length

                                If padding > 0 Then
                                    Dim paddedBlock(countSheet - 2) As Byte
                                    processBlockBytes.CopyTo(paddedBlock, padding)
                                    processBlockBytes = paddedBlock
                                ElseIf padding < 0 Then
                                    processBlockBytes = processBlockBytes.Skip(1).ToArray()
                                End If

                                decStream.Write(processBlockBytes, 0, processBlockBytes.Length)

                                idxSheet += 256
                            End While
                            decStream.Close()
                            output(TimeOfDay & " - Finished decryption." & Environment.NewLine)
                            bytes = File.ReadAllBytes(filepath & filename)
                        End If
                    End If
                End If

                If Microsoft.VisualBasic.Left(StrFromBytes(0), 4) = "DCX" Then
                    Select Case StrFromBytes(&H28)
                        Case "EDGE"
                            DCX = True
                            Dim newbytes(&H10000) As Byte
                            Dim decbytes(&H10000) As Byte
                            Dim bytes2(UIntFromBytes(&H1C) - 1) As Byte

                            Dim startOffset As UInteger = UIntFromBytes(&H14) + &H20
                            Dim numChunks As UInteger = UIntFromBytes(&H68)
                            Dim DecSize As UInteger

                            fileList = DecodeFileName(&H28) & Environment.NewLine & UInt64FromBytes(&H10) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                            For i = 0 To numChunks - 1
                                If i = numChunks - 1 Then
                                    DecSize = bytes2.Length - DecSize * i
                                Else
                                    DecSize = &H10000
                                End If

                                Array.Copy(bytes, startOffset + UIntFromBytes(&H74 + i * &H10), newbytes, 0, UIntFromBytes(&H78 + i * &H10))
                                decbytes = Decompress(newbytes)
                                Array.Copy(decbytes, 0, bytes2, &H10000 * i, DecSize)
                            Next

                            currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            File.WriteAllBytes(currFileName, bytes2)

                            File.WriteAllText(filepath & filename & ".info.txt", fileList)
                            output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)

                            bytes = bytes2
                            filepath = currFilePath
                            filename = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - filepath.Length)

                        Case "DFLT"
                            DCX = True
                            Dim startOffset As UInteger

                            If UIntFromBytes(&H14) = 76 Then
                                startOffset = UIntFromBytes(&H14) + 2
                            Else
                                startOffset = UIntFromBytes(&H14) + &H22
                            End If

                            Dim newbytes(UIntFromBytes(&H20) - 1) As Byte
                            Dim decbytes(UIntFromBytes(&H1C)) As Byte

                            fileList = DecodeFileName(&H28) & Environment.NewLine & UInt64FromBytes(&H10) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                            Array.Copy(bytes, startOffset, newbytes, 0, newbytes.Length - 2)

                            decbytes = Decompress(newbytes)

                            If IsRegulation Then
                                currFileName = filepath & filename
                            Else
                                currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                                File.WriteAllBytes(currFileName, decbytes)
                                output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)
                            End If

                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If


                            File.WriteAllText(filepath & filename & ".info.txt", fileList)


                            bytes = decbytes
                            filepath = currFilePath
                            If IsRegulation Then
                                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"
                            Else
                                filename = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - filepath.Length)
                            End If

                    End Select
                End If

                If Microsoft.VisualBasic.Left(StrFromBytes(0), 4) = "ENFL" Then
                    bigEndian = False
                    DCX = True
                    Dim startOffset As UInteger = &H12

                    Dim newbytes(UIntFromBytes(&H8) - 1) As Byte
                    Dim decbytes(UIntFromBytes(&HC)) As Byte

                    'fileList = DecodeFileName(&H28) & Environment.NewLine & UInt64FromBytes(&H10) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                    Array.Copy(bytes, startOffset, newbytes, 0, newbytes.Length - 2)

                    decbytes = Decompress(newbytes)

                    currFileName = filepath & filename & ".decompressed"
                    File.WriteAllBytes(currFileName, decbytes)
                    output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)

                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                        System.IO.Directory.CreateDirectory(currFilePath)
                    End If


                    'File.WriteAllText(filepath & filename & ".info.txt", fileList)


                    bytes = decbytes
                    filepath = currFilePath

                    filename = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - filepath.Length)
                End If

                Dim OnlyDCX = False


                Select Case Microsoft.VisualBasic.Left(StrFromBytes(0), 4)
                    Case "BHD5"

                        'DS3 BHD5 Reversing by Atvaark
                        'Credits to Atvaark and TKGP for an almost complete list of files
                        'https://github.com/Atvaark/BinderTool
                        'https://github.com/JKAnderson/UXM

                        bigEndian = False
                        If Not (UIntFromBytes(&H4) And &HFF) = &HFF Then
                            bigEndian = True
                        End If

                        fileList = "BHD5,"

                        Dim currFileSize As UInteger = 0
                        Dim currFileOffset As ULong = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim count As UInteger = 0

                        Dim idx As Integer

                        flags = UIntFromBytes(&H4)
                        numFiles = UIntFromBytes(&H10)
                        Dim startOffset As UInteger = 0

                        Dim fileidx() As String
                        Dim IsDS3 As Boolean = False

                        If flags = &H1FF Then
                            If SekiroRadio.Checked Then
                                fileidx = My.Resources.fileidx_sekiro.Replace(Chr(&HD), "").Split(Chr(&HA))
                            Else
                                fileidx = My.Resources.fileidx_ds3.Replace(Chr(&HD), "").Split(Chr(&HA))
                            End If
                            IsDS3 = True
                        Else
                            fileidx = My.Resources.fileidx.Replace(Chr(&HD), "").Split(Chr(&HA))
                        End If

                        Dim hashidx(fileidx.Length - 1) As UInteger

                        For i = 0 To fileidx.Length - 1
                            hashidx(i) = HashFileName(fileidx(i))
                        Next

                        If IsDS3 Then
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4)
                        Else
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 5)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & ".bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger

                        BinderID = ""
                        If IsDS3 Then
                            BinderID = StrFromNumBytes(&H1C, UIntFromBytes(&H18))
                        Else
                            For k = 0 To &HF
                                Dim tmpchr As Char
                                tmpchr = Chr(BDTStream.ReadByte)
                                If Not Asc(tmpchr) = 0 Then
                                    BinderID = BinderID & tmpchr
                                Else
                                    Exit For
                                End If
                            Next
                        End If

                        Dim IsSwitch As Boolean = False
                        If UIntFromBytes(&H14) = 0 And UIntFromBytes(&H18) = 32 Then
                            IsSwitch = True
                            startOffset = UIntFromBytes(&H18)
                        Else
                            startOffset = UIntFromBytes(&H14)
                        End If

                        Dim filesDict As New Dictionary(Of ULong, String)

                        fileList = fileList & BinderID & Environment.NewLine & flags & "," & Convert.ToInt32(IsSwitch) & Environment.NewLine

                        For i As UInteger = 0 To numFiles - 1

                            If IsSwitch Then
                                count = UIntFromBytes(startOffset + i * &H10)
                                bhdOffSet = UIntFromBytes(startOffset + 8 + i * 16)
                            Else
                                count = UIntFromBytes(startOffset + i * &H8)
                                bhdOffSet = UIntFromBytes(startOffset + 4 + i * 8)
                            End If

                            For j = 0 To count - 1

                                currFileSize = UIntFromBytes(bhdOffSet + &H4)

                                If bigEndian Then
                                    currFileOffset = UIntFromBytes(bhdOffSet + &HC)
                                ElseIf IsDS3 Then
                                    currFileOffset = UInt64FromBytes(bhdOffSet + &H8)
                                Else
                                    currFileOffset = UIntFromBytes(bhdOffSet + &H8)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                BDTStream.Position = currFileOffset
                                BDTStream.Read(currFileBytes, 0, currFileSize)

                                If IsDS3 Then
                                    Dim isEncrypted As Boolean = False
                                    Dim currFileSizeFinal As Long = 0
                                    Dim aesKeyOffset As Long = 0

                                    currFileSizeFinal = UInt64FromBytes(bhdOffSet + &H20)
                                    aesKeyOffset = UInt64FromBytes(bhdOffSet + &H18)
                                    If aesKeyOffset <> 0 Then
                                        isEncrypted = True
                                    End If

                                    If currFileSizeFinal = 0 Then
                                        Dim header(47) As Byte
                                        Array.Copy(currFileBytes, header, 48)

                                        If isEncrypted Then
                                            Dim aesKey(15) As Byte
                                            Array.Copy(bytes, aesKeyOffset, aesKey, 0, 16)

                                            Dim iv(15) As Byte

                                            Dim ms As New MemoryStream()
                                            Dim aes As New AesManaged() With {
                                                .Mode = CipherMode.ECB,
                                                .Padding = PaddingMode.None
                                            }

                                            Dim cs As New CryptoStream(ms, aes.CreateDecryptor(aesKey, iv), CryptoStreamMode.Write)

                                            cs.Write(header, 0, 48)

                                            Array.Copy(ms.ToArray(), header, 48)
                                            cs.Dispose()
                                        End If

                                        Dim tempBytes(3) As Byte
                                        Array.Copy(header, &H20, tempBytes, 0, 4)
                                        Array.Reverse(tempBytes)
                                        '76 -> DCX header size
                                        currFileSizeFinal = 76 + BitConverter.ToUInt32(tempBytes, 0)

                                    End If

                                    'If currFileBytes.Length >= 8 Then
                                    '    Dim firstBytes(7) As Byte
                                    '    Array.Copy(currFileBytes, firstBytes, 8)
                                    '    If UIntFromBytes(bhdOffSet) = &H57330360 Then
                                    '        Dim equal As Boolean
                                    '        For k = 0 To 7
                                    '            If firstBytes(k) = currFileBytes(k) Then
                                    '                equal = True
                                    '            Else
                                    '                equal = False
                                    '                Exit For
                                    '            End If
                                    '        Next
                                    '        If equal Then
                                    '            currFileBytes = DecryptSoundFile(currFileBytes)
                                    '        End If
                                    '    End If
                                    'End If

                                    If isEncrypted Then
                                        Dim aesKey(15) As Byte
                                        Dim numRanges As UInteger = UIntFromBytes(aesKeyOffset + &H10)
                                        Dim startOffsets(numRanges - 1) As Long
                                        Dim endOffsets(numRanges - 1) As Long

                                        For k = 0 To numRanges - 1
                                            startOffsets(k) = UInt64FromBytes(aesKeyOffset + &H14 + &H10 * k)
                                            endOffsets(k) = UInt64FromBytes(aesKeyOffset + &H1C + &H10 * k)
                                        Next

                                        Array.Copy(bytes, aesKeyOffset, aesKey, 0, 16)

                                        Dim iv(15) As Byte

                                        Dim ms As New MemoryStream()
                                        Dim aes As New AesManaged() With {
                                            .Mode = CipherMode.ECB,
                                            .Padding = PaddingMode.None
                                        }

                                        Dim cs As New CryptoStream(ms, aes.CreateDecryptor(aesKey, iv), CryptoStreamMode.Write)

                                        For k = 0 To numRanges - 1
                                            If startOffsets(k) > -1 And endOffsets(k) > -1 Then
                                                cs.Write(currFileBytes, startOffsets(k), endOffsets(k) - startOffsets(k))
                                                Array.Copy(ms.ToArray(), 0, currFileBytes, startOffsets(k), endOffsets(k) - startOffsets(k))
                                                ms.Position = 0
                                            End If
                                        Next
                                        If currFileSize > currFileSizeFinal Then
                                            ReDim Preserve currFileBytes(currFileSizeFinal - 1)
                                        End If
                                        cs.Dispose()
                                    End If


                                End If



                                'BDTStream.Position = currFileOffset

                                'For k = 0 To currFileSize - 1
                                '    currFileBytes(k) = BDTStream.ReadByte
                                'Next


                                currFileName = ""


                                If hashidx.Contains(UIntFromBytes(bhdOffSet)) Then
                                    idx = Array.IndexOf(hashidx, UIntFromBytes(bhdOffSet))

                                    currFileName = fileidx(idx)
                                    currFileName = currFileName.Replace("/", "\")


                                    'fileList += i & "," & currFileName & Environment.NewLine

                                    filesDict.Add(currFileOffset, currFileName)

                                    If IsDS3 Then
                                        currFileName = filepath & filename & ".bhd" & ".extract" & currFileName
                                    Else
                                        currFileName = filepath & filename & ".bhd5" & ".extract" & currFileName
                                    End If

                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                Else
                                    idx = -1
                                    currFileName = "NOMATCH-" & Hex(UIntFromBytes(bhdOffSet))
                                    'fileList += i & "," & currFileName & Environment.NewLine

                                    filesDict.Add(currFileOffset, currFileName)

                                    If IsDS3 Then
                                        currFileName = filepath & filename & ".bhd" & ".extract\" & currFileName
                                    Else
                                        currFileName = filepath & filename & ".bhd5" & ".extract\" & currFileName
                                    End If



                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                End If

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If


                                File.WriteAllBytes(currFileName, currFileBytes)
                                output(TimeOfDay & " - Extracted " & currFileName & Environment.NewLine)

                                If IsDS3 Then
                                    bhdOffSet += &H28
                                Else
                                    bhdOffSet += &H10
                                End If
                            Next


                        Next

                        Dim keys As List(Of ULong) = filesDict.Keys.ToList
                        keys.Sort()

                        Dim offset As ULong

                        For Each offset In keys
                            fileList += filesDict.Item(offset) & Environment.NewLine
                        Next

                        If IsDS3 Then
                            filename = filename & ".bhd"
                        Else
                            filename = filename & ".bhd5"
                        End If

                        File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)

                        'BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BHF3"
                        fileList = "BHF3,"


                        REM this assumes we'll always have between 1 and 16777215 files 
                        bigEndian = False
                        If UIntFromBytes(&H10) >= &H1000000 Then
                            bigEndian = True
                        Else
                            bigEndian = False
                        End If

                        Dim currFileSize As UInteger = 0
                        'Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim count As UInteger = 0

                        flags = UIntFromBytes(&HC)
                        numFiles = UIntFromBytes(&H10)

                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 3)

                        If Not File.Exists(filepath & filename & "bdt.bak") Then
                            File.Copy(filepath & filename & "bdt", filepath & filename & "bdt.bak")
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & "bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger = &H20

                        BinderID = StrFromBytes(&H4)
                        fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                        For i As UInteger = 0 To numFiles - 1
                            Select Case flags
                                Case &H7C, &H5C
                                    Dim currFileOffset As ULong = 0
                                    currFileSize = UIntFromBytes(bhdOffSet + &H4)
                                    currFileOffset = UInt64FromBytes(bhdOffSet + &H8)
                                    currFileID = UIntFromBytes(bhdOffSet + &H10)

                                    ReDim currFileBytes(currFileSize - 1)

                                    BDTStream.Position = currFileOffset

                                    For k = 0 To currFileSize - 1
                                        currFileBytes(k) = BDTStream.ReadByte
                                    Next


                                    currFileName = DecodeFileName(UIntFromBytes(bhdOffSet + &H14))
                                    fileList += currFileID & "," & currFileName & Environment.NewLine

                                    currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                                        System.IO.Directory.CreateDirectory(currFilePath)
                                    End If

                                    File.WriteAllBytes(currFileName, currFileBytes)

                                    bhdOffSet += &H1C
                                Case Else
                                    Dim currFileOffset As UInteger = 0
                                    currFileSize = UIntFromBytes(bhdOffSet + &H4)
                                    currFileOffset = UIntFromBytes(bhdOffSet + &H8)
                                    currFileID = UIntFromBytes(bhdOffSet + &HC)

                                    ReDim currFileBytes(currFileSize - 1)

                                    BDTStream.Position = currFileOffset

                                    For k = 0 To currFileSize - 1
                                        currFileBytes(k) = BDTStream.ReadByte
                                    Next


                                    currFileName = DecodeFileName(UIntFromBytes(bhdOffSet + &H10))
                                    fileList += currFileID & "," & currFileName & Environment.NewLine

                                    currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                                        System.IO.Directory.CreateDirectory(currFilePath)
                                    End If

                                    File.WriteAllBytes(currFileName, currFileBytes)

                                    bhdOffSet += &H18
                            End Select


                        Next
                        filename = filename & "bhd"
                        BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BHF4"
                        fileList = "BHF4,"


                        bigEndian = False

                        Dim currFileSize As ULong = 0
                        Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim unicode As Boolean = True
                        Dim count As UInteger = 0
                        Dim type As Byte = 0
                        Dim extendedHeader As Byte = 0

                        flags = UIntFromBytes(&H30)
                        unicode = flags And &HFF
                        type = (flags And &HFF00) >> 8
                        extendedHeader = flags >> 16

                        numFiles = UIntFromBytes(&HC)

                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 3)

                        If Not File.Exists(filepath & filename & "bdt.bak") Then
                            File.Copy(filepath & filename & "bdt", filepath & filename & "bdt.bak")
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & "bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger = &H40

                        BinderID = StrFromBytes(&H18)
                        fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                        For i As UInteger = 0 To numFiles - 1

                            currFileSize = UInt64FromBytes(bhdOffSet + &H8)
                            currFileOffset = UIntFromBytes(bhdOffSet + &H18)
                            currFileID = UIntFromBytes(bhdOffSet + &H1C)

                            ReDim currFileBytes(currFileSize - 1)

                            BDTStream.Position = currFileOffset

                            For k = 0 To currFileSize - 1
                                currFileBytes(k) = BDTStream.ReadByte
                            Next


                            currFileName = DecodeFileNameBND4(UIntFromBytes(bhdOffSet + &H20))
                            fileList += currFileID & "," & currFileName & Environment.NewLine

                            currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            File.WriteAllBytes(currFileName, currFileBytes)

                            bhdOffSet += &H24
                        Next
                        filename = filename & "bhd"
                        BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BND3"
                        'TODO:  DeS, c0300.anibnd, no files found?
                        'that archive is busted

                        Dim currFileSize As Long = 0
                        Dim currFileOffset As Long = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 12)
                        flags = UIntFromBytes(&HC)

                        If flags = &H74000000 Or flags = &H54000000 Or flags = &H70000000 Or flags = &H78000000 Or flags = &H7C000000 Or flags = &H5C000000 Then bigEndian = False

                        numFiles = UIntFromBytes(&H10)
                        namesEndLoc = UIntFromBytes(&H14)

                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        If numFiles = 0 Then
                            MsgBox("No files found in archive")
                            SyncLock workLock
                                work = False
                            End SyncLock
                            Exit Sub
                        End If


                        For i As UInteger = 0 To numFiles - 1
                            Select Case flags
                                Case &H70000000
                                    currFileSize = UIntFromBytes(&H24 + i * &H14)
                                    currFileOffset = UIntFromBytes(&H28 + i * &H14)
                                    currFileID = UIntFromBytes(&H2C + i * &H14)
                                    currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H74000000, &H54000000
                                    currFileSize = UIntFromBytes(&H24 + i * &H18)
                                    currFileOffset = UIntFromBytes(&H28 + i * &H18)
                                    currFileID = UIntFromBytes(&H2C + i * &H18)
                                    currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H10100
                                    currFileSize = UIntFromBytes(&H24 + i * &HC)
                                    currFileOffset = UIntFromBytes(&H28 + i * &HC)
                                    currFileID = i
                                    currFileName = i & "." & Microsoft.VisualBasic.Left(DecodeFileName(currFileOffset), 4)
                                    fileList += currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &HE010100
                                    currFileSize = UIntFromBytes(&H24 + i * &H14)
                                    currFileOffset = UIntFromBytes(&H28 + i * &H14)
                                    currFileID = UIntFromBytes(&H2C + i * &H14)
                                    currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H2E010100
                                    currFileSize = UIntFromBytes(&H24 + i * &H18)
                                    currFileOffset = UIntFromBytes(&H28 + i * &H18)
                                    currFileID = UIntFromBytes(&H2C + i * &H18)
                                    currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H78000000
                                    currFileSize = UIntFromBytes(&H24 + i * &H18)
                                    currFileOffset = UInt64FromBytes(&H28 + i * &H18)
                                    currFileID = UIntFromBytes(&H30 + i * &H18)
                                    currFileNameOffset = UIntFromBytes(&H34 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H7C000000, &H5C000000
                                    currFileSize = UIntFromBytes(&H24 + i * &H1C)
                                    currFileOffset = UInt64FromBytes(&H28 + i * &H1C)
                                    currFileID = UIntFromBytes(&H30 + i * &H1C)
                                    currFileNameOffset = UIntFromBytes(&H34 + i * &H1C)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case Else
                                    output(TimeOfDay & " - Unknown BND3 type" & Environment.NewLine)
                            End Select

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            ReDim currFileBytes(currFileSize - 1)
                            Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                            File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                        Next
                    Case "BND4"
                        Dim currFileSize As ULong = 0
                        Dim currFileOffset As ULong = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}
                        Dim currFileHash As UInteger = 0
                        Dim currFileHashIdx As UInteger = 0
                        Dim extendedHeader As Byte = 0
                        Dim unicode As Byte = 0
                        Dim type As Byte = 0
                        bigEndian = False

                        BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 4) + Microsoft.VisualBasic.Left(StrFromBytes(&H18), 8)

                        numFiles = UIntFromBytes(&HC)
                        flags = UIntFromBytes(&H30)
                        unicode = flags And &HFF
                        type = (flags And &HFF00) >> 8
                        extendedHeader = flags >> 16
                        'namesEndLoc = UIntFromBytes(&H38)


                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        If numFiles = 0 Then
                            MsgBox("No files found in archive")
                            SyncLock workLock
                                work = False
                            End SyncLock
                            Exit Sub
                        End If


                        For i As UInteger = 0 To numFiles - 1
                            Select Case type
                                Case &H74, &H54

                                    currFileSize = UInt64FromBytes(&H48 + i * &H24)
                                    currFileOffset = UIntFromBytes(&H58 + i * &H24)
                                    currFileID = UIntFromBytes(&H5C + i * &H24)
                                    currFileNameOffset = UIntFromBytes(&H60 + i * &H24)

                                    If unicode > 0 Then
                                        currFileName = DecodeFileNameBND4(currFileNameOffset)
                                    Else
                                        currFileName = DecodeFileName(currFileNameOffset)
                                    End If

                                    fileList += currFileID & "," & currFileName & Environment.NewLine

                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H20
                                    currFileSize = UInt64FromBytes(&H48 + i * &H20)
                                    currFileOffset = UIntFromBytes(&H50 + i * &H20)
                                    currFileNameOffset = UIntFromBytes(&H54 + i * &H20)

                                    If unicode > 0 Then
                                        currFileName = DecodeFileNameBND4(currFileNameOffset)
                                    Else
                                        currFileName = DecodeFileName(currFileNameOffset)
                                    End If

                                    fileList += 0 & "," & currFileName & Environment.NewLine

                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                Case Else
                                    output(TimeOfDay & " - Unknown BND4 type" & Environment.NewLine)
                            End Select

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            ReDim currFileBytes(currFileSize - 1)

                            If currFileSize > 0 Then
                                For j As ULong = 0 To currFileSize - 1
                                    currFileBytes(j) = bytes(currFileOffset + j)
                                Next
                            End If

                            'Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                            File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                        Next


                    Case "TPF"
                        Dim currFileSize As UInteger = 0
                        Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}
                        Dim currFileFlags1 As UInteger = 0
                        Dim currFileFlags2 As UInteger = 0
                        Dim currFileFormat As Byte = 0
                        Dim currFileIsCubemap As Boolean = False
                        Dim currFileMipMaps As UInteger = 0
                        Dim currFileWidth As UInt16 = 0
                        Dim currFileHeight As UInt16 = 0
                        Dim bw As BinaryWriter
                        Dim ddsHeader As DDS_HEADER

                        bigEndian = False
                        If UIntFromBytes(&H8) >= &H1000000 Then
                            bigEndian = True
                        Else
                            bigEndian = False
                        End If

                        flags = UIntFromBytes(&HC)

                        If flags = &H2010200 Or flags = &H2010000 Or flags = &H2020200 Or flags = &H2030200 Then
                            ' Dark Souls/Demon's Souls (headerless DDS)

                            BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                            numFiles = UIntFromBytes(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1

                                writeOffset = 0
                                currFileOffset = UIntFromBytes(&H10 + i * &H20)
                                currFileSize = UIntFromBytes(&H14 + i * &H20)
                                currFileFormat = bytes(&H18 + i * &H20)
                                currFileIsCubemap = bytes(&H19 + i * &H20)
                                currFileMipMaps = bytes(&H1A + i * &H20)
                                currFileWidth = UInt16FromBytes(&H1C + i * &H20)
                                currFileHeight = UInt16FromBytes(&H1E + i * &H20)
                                currFileFlags1 = UIntFromBytes(&H20 + i * &H20)
                                currFileFlags2 = UIntFromBytes(&H24 + i * &H20)
                                currFileNameOffset = UIntFromBytes(&H28 + i * &H20)
                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
                                fileList += currFileFormat & "," & Convert.ToInt32(currFileIsCubemap) & "," & currFileMipMaps & "," & currFileFlags1 & "," & currFileFlags2 & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If


                                ddsHeader = MakeDDSHeader(currFileFormat, currFileIsCubemap, currFileMipMaps, currFileHeight, currFileWidth)

                                bw = New BinaryWriter(File.Open(currFilePath & currFileName, FileMode.Create))

                                bw.Write(&H20534444)
                                bw.Write(ddsHeader.dwSize)
                                bw.Write(ddsHeader.dwFlags)
                                bw.Write(ddsHeader.dwHeight)
                                bw.Write(ddsHeader.dwWidth)
                                bw.Write(ddsHeader.dwPitchOrLinearSize)
                                bw.Write(ddsHeader.dwDepth)
                                bw.Write(ddsHeader.dwMipMapCount)
                                bw.Seek(&H4C, 0)
                                bw.Write(ddsHeader.ddspf.dwSize)
                                bw.Write(ddsHeader.ddspf.dwFlags)
                                bw.Write(ddsHeader.ddspf.dwFourCC)
                                bw.Write(ddsHeader.ddspf.dwRGBBitCount)
                                bw.Write(ddsHeader.ddspf.dwRBitMask)
                                bw.Write(ddsHeader.ddspf.dwGBitMask)
                                bw.Write(ddsHeader.ddspf.dwBBitMask)
                                bw.Write(ddsHeader.ddspf.dwABitMask)
                                bw.Write(ddsHeader.dwCaps)
                                bw.Write(ddsHeader.dwCaps2)
                                bw.Seek(&H80, 0)

                                If (currFileFormat = 10 Or currFileFormat = 9) Then
                                    ddsWidth = currFileWidth
                                    If currFileIsCubemap Then
                                        ReDim inBytes(ddsHeader.dwPitchOrLinearSize * (4 / 3) - 1)
                                        ReDim outBytes(ddsHeader.dwPitchOrLinearSize - 1)
                                        ReDim currFileBytes(ddsHeader.dwPitchOrLinearSize * 6 - 1)
                                        'Deswizzle cube faces separately
                                        For j As UInteger = 0 To 5
                                            Array.Copy(bytes, CType(currFileOffset + j * ddsHeader.dwPitchOrLinearSize * (4 / 3), UInteger), inBytes, 0, CType(ddsHeader.dwPitchOrLinearSize * (4 / 3), UInteger))
                                            DeswizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                            writeOffset = 0
                                            Array.Copy(outBytes, 0, currFileBytes, j * ddsHeader.dwPitchOrLinearSize, ddsHeader.dwPitchOrLinearSize)
                                        Next
                                        bw.Write(currFileBytes)
                                    Else
                                        ReDim inBytes(currFileSize - 1)
                                        ReDim outBytes(currFileSize * (3 / 4) - 1)
                                        Array.Copy(bytes, currFileOffset, inBytes, 0, currFileSize)
                                        DeswizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                        bw.Write(outBytes)
                                    End If
                                ElseIf currFileIsCubemap Then
                                    'This one has meme bytes
                                    ReDim currFileBytes(currFileSize - 1)
                                    Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                                    RemoveMemeBytes(currFileBytes)
                                    bw.Write(currFileBytes)
                                Else
                                    ReDim currFileBytes(currFileSize - 1)
                                    Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                                    bw.Write(currFileBytes)
                                End If

                                bw.Close()
                            Next

                        ElseIf flags = &H1030200 Then
                            ' Dark Souls (headerless DDS)

                            BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                            numFiles = UIntFromBytes(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1

                                writeOffset = 0
                                currFileOffset = UIntFromBytes(&H10 + i * &H1C)
                                currFileSize = UIntFromBytes(&H14 + i * &H1C)
                                currFileFormat = bytes(&H18 + i * &H1C)
                                currFileIsCubemap = bytes(&H19 + i * &H1C)
                                currFileMipMaps = bytes(&H1A + i * &H1C)
                                currFileWidth = UInt16FromBytes(&H1C + i * &H1C)
                                currFileHeight = UInt16FromBytes(&H1E + i * &H1C)
                                currFileFlags1 = UIntFromBytes(&H20 + i * &H1C)
                                currFileNameOffset = UIntFromBytes(&H24 + i * &H1C)
                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
                                fileList += currFileFormat & "," & Convert.ToInt32(currFileIsCubemap) & "," & currFileMipMaps & "," & currFileFlags1 & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If


                                ddsHeader = MakeDDSHeader(currFileFormat, currFileIsCubemap, currFileMipMaps, currFileHeight, currFileWidth)

                                bw = New BinaryWriter(File.Open(currFilePath & currFileName, FileMode.Create))

                                bw.Write(&H20534444)
                                bw.Write(ddsHeader.dwSize)
                                bw.Write(ddsHeader.dwFlags)
                                bw.Write(ddsHeader.dwHeight)
                                bw.Write(ddsHeader.dwWidth)
                                bw.Write(ddsHeader.dwPitchOrLinearSize)
                                bw.Write(ddsHeader.dwDepth)
                                bw.Write(ddsHeader.dwMipMapCount)
                                bw.Seek(&H4C, 0)
                                bw.Write(ddsHeader.ddspf.dwSize)
                                bw.Write(ddsHeader.ddspf.dwFlags)
                                bw.Write(ddsHeader.ddspf.dwFourCC)
                                bw.Write(ddsHeader.ddspf.dwRGBBitCount)
                                bw.Write(ddsHeader.ddspf.dwRBitMask)
                                bw.Write(ddsHeader.ddspf.dwGBitMask)
                                bw.Write(ddsHeader.ddspf.dwBBitMask)
                                bw.Write(ddsHeader.ddspf.dwABitMask)
                                bw.Write(ddsHeader.dwCaps)
                                bw.Write(ddsHeader.dwCaps2)
                                bw.Seek(&H80, 0)

                                If (currFileFormat = 10 Or currFileFormat = 9) Then
                                    ddsWidth = currFileWidth
                                    If currFileIsCubemap Then
                                        ReDim inBytes(ddsHeader.dwPitchOrLinearSize * (4 / 3) - 1)
                                        ReDim outBytes(ddsHeader.dwPitchOrLinearSize - 1)
                                        ReDim currFileBytes(ddsHeader.dwPitchOrLinearSize * 6 - 1)
                                        'Deswizzle cube faces separately
                                        For j As UInteger = 0 To 5
                                            Array.Copy(bytes, CType(currFileOffset + j * ddsHeader.dwPitchOrLinearSize * (4 / 3), UInteger), inBytes, 0, CType(ddsHeader.dwPitchOrLinearSize * (4 / 3), UInteger))
                                            DeswizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                            writeOffset = 0
                                            Array.Copy(outBytes, 0, currFileBytes, j * ddsHeader.dwPitchOrLinearSize, ddsHeader.dwPitchOrLinearSize)
                                        Next
                                        bw.Write(currFileBytes)
                                    Else
                                        ReDim inBytes(currFileSize - 1)
                                        ReDim outBytes(currFileSize * (3 / 4) - 1)
                                        Array.Copy(bytes, currFileOffset, inBytes, 0, currFileSize)
                                        DeswizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                        bw.Write(outBytes)
                                    End If
                                ElseIf currFileIsCubemap Then
                                    'This one has meme bytes
                                    ReDim currFileBytes(currFileSize - 1)
                                    Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                                    If currFileSize >= &H41B8 Then
                                        RemoveMemeBytes(currFileBytes)
                                    End If
                                    bw.Write(currFileBytes)
                                End If

                                bw.Close()
                            Next

                        ElseIf flags = &H20300 Or flags = &H20304 Then
                            ' Dark Souls

                            BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                            numFiles = UIntFromBytes(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = UIntFromBytes(&H10 + i * &H14)
                                currFileSize = UIntFromBytes(&H14 + i * &H14)
                                currFileFlags1 = UIntFromBytes(&H18 + i * &H14)
                                currFileNameOffset = UIntFromBytes(&H1C + i * &H14)
                                currFileFlags2 = UIntFromBytes(&H20 + i * &H14)
                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
                                fileList += currFileFlags1 & "," & currFileFlags2 & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                                File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                            Next
                        ElseIf flags = &H10300 Then
                            ' Dark Souls III

                            BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                            numFiles = UIntFromBytes(&H8)
                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = UIntFromBytes(&H10 + i * &H14)
                                currFileSize = UIntFromBytes(&H14 + i * &H14)
                                currFileFormat = bytes(&H18 + i * &H14)
                                currFileIsCubemap = bytes(&H19 + i * &H14)
                                currFileMipMaps = bytes(&H1A + i * &H14)
                                currFileNameOffset = UIntFromBytes(&H1C + i * &H14)
                                currFileFlags1 = UIntFromBytes(&H20 + i * &H14)

                                currFileName = DecodeFileNameBND4(currFileNameOffset) & ".dds"
                                fileList += currFileFormat & "," & Convert.ToInt32(currFileIsCubemap) & "," & currFileMipMaps & "," & currFileFlags1 & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                                File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                            Next
                        ElseIf flags = &H10304 Then
                            ' Dark Souls III (headerless DDS)

                            Dim currFileDxgiFormat As UInteger = 0
                            Dim currFileArraySize As UInteger = 0
                            Dim ddsDx10Header As DDS_HEADER_DXT10
                            BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                            numFiles = UIntFromBytes(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = UIntFromBytes(&H10 + i * &H24)
                                currFileSize = UIntFromBytes(&H14 + i * &H24)
                                currFileFormat = bytes(&H18 + i * &H24)
                                currFileIsCubemap = bytes(&H19 + i * &H24)
                                currFileMipMaps = bytes(&H1A + i * &H24)
                                currFileWidth = UInt16FromBytes(&H1C + i * &H24)
                                currFileHeight = UInt16FromBytes(&H1E + i * &H24)
                                currFileArraySize = UIntFromBytes(&H20 + i * &H24)
                                'unk always 0D?
                                currFileNameOffset = UIntFromBytes(&H28 + i * &H24)
                                currFileDxgiFormat = UIntFromBytes(&H30 + i * &H24)
                                currFileName = DecodeFileNameBND4(currFileNameOffset) & ".dds"
                                fileList += currFileFormat & "," & Convert.ToInt32(currFileIsCubemap) & "," & currFileMipMaps & "," & currFileArraySize & "," & currFileDxgiFormat & "," & currFileName & Environment.NewLine

                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If

                                ddsHeader = MakeDDSHeaderPS4(currFileFormat, currFileIsCubemap, currFileMipMaps, currFileHeight, currFileWidth)
                                ddsDx10Header = MakeDDSDX10Header(currFileDxgiFormat, currFileIsCubemap)

                                bw = New BinaryWriter(File.Open(currFilePath & currFileName, FileMode.Create))

                                bw.Write(&H20534444)
                                bw.Write(ddsHeader.dwSize)
                                bw.Write(ddsHeader.dwFlags)
                                bw.Write(ddsHeader.dwHeight)
                                bw.Write(ddsHeader.dwWidth)
                                bw.Write(ddsHeader.dwPitchOrLinearSize)
                                bw.Write(ddsHeader.dwDepth)
                                bw.Write(ddsHeader.dwMipMapCount)
                                bw.Seek(&H4C, 0)
                                bw.Write(ddsHeader.ddspf.dwSize)
                                bw.Write(ddsHeader.ddspf.dwFlags)
                                bw.Write(ddsHeader.ddspf.dwFourCC)
                                bw.Write(ddsHeader.ddspf.dwRGBBitCount)
                                bw.Write(ddsHeader.ddspf.dwRBitMask)
                                bw.Write(ddsHeader.ddspf.dwGBitMask)
                                bw.Write(ddsHeader.ddspf.dwBBitMask)
                                bw.Write(ddsHeader.ddspf.dwABitMask)
                                bw.Write(ddsHeader.dwCaps)
                                bw.Write(ddsHeader.dwCaps2)
                                bw.Seek(&H80, 0)

                                If currFileFormat <> 22 And currFileFormat <> 105 Then
                                    bw.Write(ddsDx10Header.dxgiFormat)
                                    bw.Write(ddsDx10Header.resourceDimension)
                                    bw.Write(ddsDx10Header.miscFlag)
                                    bw.Write(ddsDx10Header.arraySize)
                                    bw.Write(ddsDx10Header.miscFlags2)
                                End If


                                ReDim currFileBytes(currFileSize - 1)
                                Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)


                                Dim paddedWidth As UInteger = 0
                                Dim paddedHeight As UInteger = 0
                                Dim paddedSize As UInteger = 0
                                Dim Width = currFileWidth
                                Dim Height = currFileHeight
                                Dim BlockSize As UInteger = 0
                                Dim copyOffset = currFileOffset

                                Select Case (currFileFormat)
                                    Case 22, 25
                                        BlockSize = 8
                                    Case 105
                                        BlockSize = 4
                                    Case 0, 1, 103, 108, 109
                                        BlockSize = 8
                                    Case 5, 100, 102, 106, 107, 110
                                        BlockSize = 16
                                    Case Else
                                        MsgBox("Format " & currFileFormat & " in file " & currFileName & " not implemented. Aborting :)")
                                        SyncLock workLock
                                            work = False
                                        End SyncLock
                                        bw.Close()
                                        File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)
                                        Return
                                End Select

                                If currFileIsCubemap Then
                                    For j As UInteger = 0 To currFileArraySize - 1
                                        Width = currFileWidth
                                        Height = currFileHeight
                                        If currFileFormat = 22 Then
                                            paddedWidth = Math.Ceiling(Width / 8) * 8
                                            paddedHeight = Math.Ceiling(Height / 8) * 8
                                            paddedSize = paddedWidth * paddedHeight * BlockSize
                                        Else
                                            paddedWidth = Math.Ceiling(Width / 32) * 32
                                            paddedHeight = Math.Ceiling(Height / 32) * 32
                                            paddedSize = Math.Ceiling(paddedWidth / 4) * Math.Ceiling(paddedHeight / 4) * BlockSize
                                        End If

                                        copyOffset = currFileOffset + paddedSize * j
                                        For k As UInteger = 0 To ddsHeader.dwMipMapCount - 1
                                            If k > 0 Then
                                                If currFileFormat = 22 Then
                                                    paddedWidth = Math.Ceiling(Width / 8) * 8
                                                    paddedHeight = Math.Ceiling(Height / 8) * 8
                                                    paddedSize = paddedWidth * paddedHeight * BlockSize
                                                Else
                                                    paddedWidth = Math.Ceiling(Width / 32) * 32
                                                    paddedHeight = Math.Ceiling(Height / 32) * 32
                                                    paddedSize = Math.Ceiling(paddedWidth / 4) * Math.Ceiling(paddedHeight / 4) * BlockSize
                                                End If


                                                copyOffset += paddedSize * j
                                            End If

                                            ReDim outBytes(paddedSize - 1)
                                            ReDim inBytes(paddedSize - 1)
                                            Array.Copy(bytes, copyOffset, inBytes, 0, paddedSize)
                                            ddsWidth = paddedWidth
                                            DeswizzleDDSBytesPS4(paddedWidth, paddedHeight, currFileFormat)

                                            If currFileFormat = 22 Then
                                                For l As UInteger = 0 To Height - 1
                                                    bw.Write(outBytes, l * paddedWidth * BlockSize, Width * BlockSize)
                                                Next
                                            Else
                                                For l As UInteger = 0 To Math.Ceiling(Height / 4) - 1
                                                    bw.Write(outBytes, CType(l * Math.Ceiling(paddedWidth / 4) * BlockSize, UInteger), CType(Math.Ceiling(Width / 4) * BlockSize, UInteger))
                                                Next
                                            End If

                                            copyOffset += (currFileArraySize - j) * paddedSize + paddedSize * 2

                                            If Width > 1 Then
                                                Width /= 2
                                            End If
                                            If Height > 1 Then
                                                Height /= 2
                                            End If
                                        Next
                                    Next
                                Else
                                    For j As UInteger = 0 To ddsHeader.dwMipMapCount - 1

                                        If currFileFormat = 105 Then
                                            paddedWidth = Width
                                            paddedHeight = Height
                                            paddedSize = paddedWidth * paddedHeight * BlockSize
                                        Else
                                            paddedWidth = Math.Ceiling(Width / 32) * 32
                                            paddedHeight = Math.Ceiling(Height / 32) * 32
                                            paddedSize = Math.Ceiling(paddedWidth / 4) * Math.Ceiling(paddedHeight / 4) * BlockSize
                                        End If

                                        ddsWidth = paddedWidth
                                        ReDim inBytes(paddedSize - 1)
                                        ReDim outBytes(paddedSize - 1)
                                        Array.Copy(bytes, copyOffset, inBytes, 0, paddedSize)
                                        DeswizzleDDSBytesPS4(paddedWidth, paddedHeight, currFileFormat)

                                        If currFileFormat = 105 Then
                                            bw.Write(outBytes)
                                        Else
                                            For k As UInteger = 0 To Math.Ceiling(Height / 4) - 1
                                                bw.Write(outBytes, CType(k * Math.Ceiling(paddedWidth / 4) * BlockSize, UInteger), CType(Math.Ceiling(Width / 4) * BlockSize, UInteger))
                                            Next
                                        End If

                                        copyOffset += paddedSize
                                        If Width > 1 Then
                                            Width /= 2
                                        End If
                                        If Height > 1 Then
                                            Height /= 2
                                        End If
                                    Next
                                End If


                                bw.Close()
                            Next
                        Else
                            output(TimeOfDay & " - Unknown TPF format" & Environment.NewLine)
                        End If

                    Case Else
                        OnlyDCX = True

                End Select

                If Not OnlyDCX Then
                    File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)
                    output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)
                End If


            Next



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)
            output(TimeOfDay & " - Unhandled exception - " & ex.Message & ex.StackTrace & Environment.NewLine)
        End Try

        SyncLock workLock
            work = False
        End SyncLock
        'txtInfo.Text += TimeOfDay & " - " & filename & " extracted." & Environment.NewLine
    End Sub
    Private Sub BtnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        trdWorker = New Thread(AddressOf Rebuild) With {
            .IsBackground = True
        }
        trdWorker.Start()
    End Sub
    Private Sub Rebuild()
        'TODO:  Confirm endian before each rebuild.

        'TODO:  List of non-DCXs that don't rebuild byte-perfect
        '   DeS, facegen.tpf
        '   DeS, i7006.tpf
        '   DeS, m07_9990.tpf
        '   DaS, m10_9999.tpf
        SyncLock workLock
            work = True
        End SyncLock

        Try

            For Each bndfile In txtBNDfile.Lines
                bigEndian = True

                Dim DCX As Boolean = False
                Dim OnlyDCX As Boolean = False
                Dim IsRegulation As Boolean = False
                Dim IsSl2 As Boolean = False

                Dim currFileSize As UInteger = 0
                Dim currFileOffset As Long = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileName As String = ""
                Dim currFilePath As String = ""
                Dim currFileBytes() As Byte = {}
                Dim currFileID As UInteger = 0
                Dim namesEndLoc As UInteger = 0
                Dim fileList As String() = {""}
                Dim BinderID As String = ""
                Dim flags As UInteger = 0
                Dim numFiles As UInteger = 0
                Dim tmpbytes() As Byte
                Dim dcxBytes() As Byte

                Dim padding As UInteger = 0

                filepath = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, "\"))
                filename = Microsoft.VisualBasic.Right(bndfile, bndfile.Length - filepath.Length)
                DCX = (Microsoft.VisualBasic.Right(filename, 4).ToLower = ".dcx")

                If Microsoft.VisualBasic.Right(filename, 3) = "bhd" Then
                    bytes = File.ReadAllBytes(filepath & filename)
                    Dim firstBytes As UInteger = UIntFromBytes(&H0)
                    If archiveDict.ContainsKey(firstBytes) Then
                        If archiveDict(firstBytes) = "Data0" Then
                            IsRegulation = True
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"
                            MsgBox("Using a modified Data0.bdt (regulation file) online will get you banned. Proceed at your own risk.", MessageBoxIcon.Warning)
                        End If
                    End If
                End If

                If DCX = True Then
                    dcxBytes = File.ReadAllBytes(filepath & filename)
                    filename = filename.Substring(0, filename.Length - 4)
                End If
                Try
                    output(TimeOfDay & " - Processing filelist.txt..." & Environment.NewLine)
                    'If Not DCX Then
                    fileList = File.ReadAllLines(filepath & filename & ".extract\" & "fileList.txt")
                    'Else
                    'fileList = File.ReadAllLines(filepath & filename & ".info.txt")
                    'End If
                Catch ex As DirectoryNotFoundException
                    OnlyDCX = True
                Catch ex As Exception
                    MsgBox(ex.Message, MessageBoxIcon.Error)
                    SyncLock workLock
                        work = False
                    End SyncLock
                    Return
                End Try

                If OnlyDCX = False Then
                    If IsRegulation Then
                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                    End If
                    If Not File.Exists(filepath & filename & ".bak") Then
                        bytes = File.ReadAllBytes(filepath & filename)
                        File.WriteAllBytes(filepath & filename & ".bak", bytes)
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                        output(TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine)
                    Else
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                        output(TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine)
                    End If

                    Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
                        Case "BHD5"
                            BinderID = fileList(0).Split(",")(1)

                            If fileList(1).Split(",").Length < 2 Then
                                MsgBox("filelist.txt incompatible. Please extract once more before rebuilding.", MessageBoxIcon.Error)
                                SyncLock workLock
                                    work = False
                                End SyncLock
                                Return
                            End If
                            output(TimeOfDay & " - Beginning BHD5 rebuild." & Environment.NewLine)

                            Dim IsSwitch As Boolean = fileList(1).Split(",")(1)

                            flags = fileList(1).Split(",")(0)
                            numFiles = fileList.Length - 2
                            If flags = 0 Then
                                bigEndian = True
                            Else
                                bigEndian = False
                            End If

                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, ".")) & "bdt"

                            If Not File.Exists(BDTFilename & ".bak") Then
                                File.Copy(BDTFilename, BDTFilename & ".bak")
                                'txtinfo.text += timeofday & " - " & filename & ".bdt.bak created." & environment.newline
                                output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                            Else
                                'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                                output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                            End If

                            Dim IsDS3 As Boolean = False

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            Dim bdtoffset As ULong = 0

                            'Dim bins(fileList.Length - 2) As UInteger
                            'Dim currBin As UInteger = 0
                            'Dim totBin As UInteger = 0

                            Dim bucketEntryLength = &H10
                            Dim bucketLength = &H8

                            'For i = 0 To fileList.Length - 3
                            '    currBin = fileList(i + 2).Split(",")(0)
                            '    bins(currBin) += 1
                            'Next
                            'totBin = Val(fileList(numFiles + 1).Split(",")(0)) + 1

                            Dim bucketOffset As UInteger = 0
                            Dim bucketEntryOffset As UInteger = 0
                            Dim startOffset As UInteger

                            Select Case flags
                                Case &H1FF
                                    IsDS3 = True
                                    startOffset = &H1C + BinderID.Length
                                    bucketEntryLength = &H28

                                Case Else
                                    BDTStream.Position = 0
                                    WriteBytes(BDTStream, StrToBytes(BinderID))
                                    BDTStream.Position = &H10

                                    bdtoffset = &H10

                                    If IsSwitch Then
                                        startOffset = &H20
                                        bucketLength = &H10
                                    Else
                                        startOffset = &H18
                                    End If

                            End Select

                            bucketOffset = startOffset
                            ReDim bytes(startOffset - 1)

                            If IsDS3 Then
                                UIntToBytes(BinderID.Length, &H18)
                                StrToBytes(BinderID, &H1C)
                            End If

                            Dim groupCount As UInteger


                            For i As UInteger = numFiles \ 7 To 100000
                                Dim noPrime = False
                                For j As UInteger = 2 To i - 1
                                    If i Mod j = 0 Or i = 2 Then
                                        noPrime = True
                                        Exit For
                                    End If
                                Next
                                If noPrime = False And i > 1 Then
                                    groupCount = i
                                    Exit For
                                End If
                            Next

                            'MsgBox("groupcount: " & groupCount)

                            StrToBytes("BHD5", 0)
                            UIntToBytes(flags, &H4)
                            UIntToBytes(1, &H8)
                            'total file size, &HC
                            UIntToBytes(groupCount, &H10)
                            If IsSwitch Then
                                UIntToBytes(startOffset, &H18)
                            Else
                                UIntToBytes(startOffset, &H14)
                            End If

                            'idxOffset = startOffset + groupCount * bucketLength

                            ReDim Preserve bytes((startOffset - 1) + groupCount * bucketLength)

                            Dim hashLists(groupCount) As List(Of pathHash)


                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i) = New List(Of pathHash)
                            Next

                            Dim hashGroups As New List(Of hashGroup)
                            Dim pathHashes As New List(Of pathHash)

                            'Dim IsCompressed As Boolean = False
                            'Dim firstChars(2) As Byte
                            'Dim saltedHashes As Dictionary(Of UInteger, saltedHash) = New Dictionary(Of UInteger, saltedHash)
                            'Dim count As UInteger = 0

                            For i As UInteger = 0 To numFiles - 1
                                Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                Dim pathHash As pathHash = New pathHash()
                                If internalFileName(0) <> "\" Then
                                    internalFileName = "\" & internalFileName
                                End If
                                Dim hash As UInteger = HashFileName(internalFileName.Replace("\", "/"))

                                Dim fStream As New IO.FileStream(filepath & filename & ".extract" & internalFileName, IO.FileMode.Open)

                                pathHash.hash = hash
                                pathHash.filesize = fStream.Length
                                pathHash.bdtoffset = bdtoffset
                                pathHash.idx = i
                                Dim group As UInteger = hash Mod groupCount
                                hashLists(group).Add(pathHash)

                                'Dim saltedHash As saltedHash

                                'fStream.Read(firstChars, 0, 3)
                                'fStream.Position = 0

                                'Calculate salted SHA256 hashes
                                'Commented out for the time being, since it was only used for encrypted files
                                '...and we don't encrypt
                                '...and the pattern of which ranges to hash I thought I recognized is wrong

                                'If IsDS3 Then
                                '    If firstChars(0) = &H44 And firstChars(1) = &H43 And firstChars(2) = &H58 Then
                                '        IsCompressed = True
                                '    Else
                                '        saltedHash.idx = count
                                '        saltedHash.rangeCount = 3
                                '        saltedHash.ranges = New List(Of bhd5Range)
                                '        Dim range As bhd5Range
                                '        range.startOffset = 0
                                '        If fStream.Length > &H100 Then
                                '            range.endOffset = &H100
                                '        Else
                                '            range.endOffset = fStream.Length - 1
                                '        End If
                                '        saltedHash.ranges.Add(range)
                                '        If fStream.Length > &H200 Then
                                '            range.startOffset = &H200
                                '            If fStream.Length > &H280 Then
                                '                range.endOffset = &H280
                                '                saltedHash.ranges.Add(range)
                                '                If fStream.Length > &H19000 Then
                                '                    range.startOffset = &H19000
                                '                    If fStream.Length > &H19080 Then
                                '                        range.endOffset = &H19080
                                '                    Else
                                '                        range.endOffset = fStream.Length - 1
                                '                    End If
                                '                    saltedHash.ranges.Add(range)
                                '                Else
                                '                    range.startOffset = -1
                                '                    range.endOffset = -1
                                '                    saltedHash.ranges.Add(range)
                                '                End If
                                '            Else
                                '                range.endOffset = fStream.Length - 1
                                '                saltedHash.ranges.Add(range)
                                '                range.startOffset = -1
                                '                range.endOffset = -1
                                '                saltedHash.ranges.Add(range)
                                '            End If
                                '        Else
                                '            range.startOffset = -1
                                '            range.endOffset = -1
                                '            saltedHash.ranges.Add(range)
                                '            saltedHash.ranges.Add(range)
                                '        End If

                                '        Dim byteCount As UInteger = saltedHash.ranges(0).endOffset
                                '        Dim tempLength As UInteger = 0
                                '        Dim byteBuffer(byteCount - 1) As Byte
                                '        Dim hashBytes(byteCount - 1) As Byte
                                '        fStream.Read(byteBuffer, 0, byteCount - 1)
                                '        Array.Copy(byteBuffer, 0, hashBytes, 0, byteCount - 1)

                                '        If saltedHash.ranges(1).startOffset > -1 Then
                                '            byteCount = saltedHash.ranges(1).endOffset - saltedHash.ranges(1).startOffset
                                '            tempLength = hashBytes.Length
                                '            ReDim Preserve hashBytes(hashBytes.Length + byteCount - 1)
                                '            fStream.Position = saltedHash.ranges(1).startOffset
                                '            fStream.Read(hashBytes, tempLength, byteCount)

                                '            If saltedHash.ranges(2).startOffset > -1 Then
                                '                byteCount = saltedHash.ranges(2).endOffset - saltedHash.ranges(2).startOffset
                                '                tempLength = hashBytes.Length
                                '                ReDim Preserve hashBytes(hashBytes.Length + byteCount - 1)
                                '                fStream.Position = saltedHash.ranges(2).startOffset
                                '                fStream.Read(hashBytes, tempLength, byteCount)
                                '            End If
                                '        End If

                                '        tempLength = hashBytes.Length
                                '        ReDim Preserve hashBytes(hashBytes.Length + BinderID.Length - 1)

                                '        Array.Copy(StrToBytes(BinderID), 0, hashBytes, tempLength, BinderID.Length)

                                '        Dim SHA As SHA256 = SHA256Managed.Create()

                                '        saltedHash.shaHash = SHA.ComputeHash(hashBytes)

                                '        SHA.Dispose()

                                '        saltedHashes.Add(i, saltedHash)

                                '        fStream.Position = 0
                                '        count += 1
                                '    End If


                                'End If


                                For j = 0 To fStream.Length - 1
                                    BDTStream.WriteByte(fStream.ReadByte)
                                Next

                                bdtoffset = BDTStream.Position
                                If bdtoffset Mod &H10 > 0 Then
                                    padding = &H10 - (bdtoffset Mod &H10)
                                Else
                                    padding = 0
                                End If
                                bdtoffset += padding

                                BDTStream.Position = bdtoffset

                                fStream.Dispose()
                                output(TimeOfDay & " - Added " & internalFileName & Environment.NewLine)
                            Next

                            bucketEntryOffset = startOffset + groupCount * bucketLength
                            'Dim saltedHashesOffset As ULong = bucketEntryOffset + numFiles * bucketEntryLength
                            'ReDim Preserve bytes(bytes.Length + numFiles * bucketEntryLength - 1 + saltedHashes.Count * &H54)
                            ReDim Preserve bytes(bytes.Length + numFiles * bucketEntryLength - 1)

                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i).Sort(Function(x, y) x.bdtoffset.CompareTo(y.bdtoffset))
                                UIntToBytes(hashLists(i).Count, startOffset + i * bucketLength)
                                If IsSwitch Then
                                    UIntToBytes(&H1, startOffset + 4 + i * bucketLength)
                                    UIntToBytes(bucketEntryOffset, startOffset + 8 + i * bucketLength)
                                Else
                                    UIntToBytes(bucketEntryOffset, startOffset + 4 + i * bucketLength)
                                End If
                                'idxOffset += hashLists(i).Count * bucketEntryLength

                                For Each pathHash As pathHash In hashLists(i)
                                    UIntToBytes(pathHash.hash, bucketEntryOffset)
                                    UIntToBytes(pathHash.filesize, bucketEntryOffset + &H4)
                                    If bigEndian Then
                                        UIntToBytes(pathHash.bdtoffset, bucketEntryOffset + &HC)
                                    ElseIf IsDS3 Then
                                        UInt64ToBytes(pathHash.bdtoffset, bucketEntryOffset + &H8)
                                        '   Writing salted hash entry to file
                                        'If saltedHashes.ContainsKey(pathHash.idx) Then
                                        '    Dim saltedHash As saltedHash = saltedHashes(pathHash.idx)
                                        '    Dim tempOffset As ULong = saltedHashesOffset + saltedHash.idx * &H54
                                        '    UInt64ToBytes(tempOffset, bucketEntryOffset + &H10)

                                        '    InsBytes(saltedHash.shaHash, tempOffset)
                                        '    UIntToBytes(3, tempOffset + &H20)
                                        '    UInt64ToBytes(saltedHash.ranges(0).startOffset, tempOffset + &H24)
                                        '    UInt64ToBytes(saltedHash.ranges(0).endOffset, tempOffset + &H2C)
                                        '    UInt64ToBytes(saltedHash.ranges(1).startOffset, tempOffset + &H34)
                                        '    UInt64ToBytes(saltedHash.ranges(1).endOffset, tempOffset + &H3C)
                                        '    UInt64ToBytes(saltedHash.ranges(2).startOffset, tempOffset + &H44)
                                        '    UInt64ToBytes(saltedHash.ranges(2).endOffset, tempOffset + &H4C)
                                        'End If
                                    Else
                                        UIntToBytes(pathHash.bdtoffset, bucketEntryOffset + &H8)
                                    End If
                                    pathHashes.Add(pathHash)
                                    bucketEntryOffset += bucketEntryLength
                                Next
                            Next

                            UIntToBytes(bytes.Length, &HC)

                            BDTStream.Dispose()

                        Case "BHF3"
                            BinderID = fileList(0).Split(",")(1)
                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            Dim currNameOffset As UInteger = 0

                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, bndfile.Length - 3) & "bdt"

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            BDTStream.Position = 0
                            WriteBytes(BDTStream, StrToBytes("BDF3" & BinderID))
                            BDTStream.Position = &H10

                            ReDim bytes(&H1F)

                            Dim bdtoffset As ULong = &H10
                            Dim unk As UInteger = 0

                            StrToBytes("BHF3" & BinderID, 0)

                            If flags = &H74 Or flags = &H54 Or flags = &H7C Or flags = &H5C Then
                                bigEndian = False
                                unk = &H40
                            Else
                                unk = &H2000000
                            End If

                            UIntToBytes(flags, &HC)
                            UIntToBytes(numFiles, &H10)

                            Dim elemLength As UInteger = &H18

                            If flags = &H7C Or flags = &H5C Then elemLength = &H1C

                            ReDim Preserve bytes(&H1F + numFiles * elemLength)


                            Dim idxOffset As UInteger
                            idxOffset = &H20


                            For i = 0 To numFiles - 1
                                currFileID = fileList(i + 2).Split(",")(0)
                                currFileName = fileList(i + 2).Split(",")(1)
                                currNameOffset = bytes.Length

                                Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                                UIntToBytes(unk, idxOffset + i * elemLength)
                                UIntToBytes(fStream.Length, idxOffset + &H4 + i * elemLength)
                                If flags = &H7C Or flags = &H5C Then
                                    UInt64ToBytes(bdtoffset, idxOffset + &H8 + i * elemLength)
                                    UIntToBytes(currFileID, idxOffset + &H10 + i * elemLength)
                                    UIntToBytes(currNameOffset, idxOffset + &H14 + i * elemLength)
                                    UIntToBytes(fStream.Length, idxOffset + &H18 + i * elemLength)
                                Else
                                    UIntToBytes(bdtoffset, idxOffset + &H8 + i * elemLength)
                                    UIntToBytes(currFileID, idxOffset + &HC + i * elemLength)
                                    UIntToBytes(currNameOffset, idxOffset + &H10 + i * elemLength)
                                    UIntToBytes(fStream.Length, idxOffset + &H14 + i * elemLength)
                                End If

                                ReDim Preserve bytes(bytes.Length + currFileName.Length)

                                EncodeFileName(currFileName, currNameOffset)

                                For j = 0 To fStream.Length - 1
                                    BDTStream.WriteByte(fStream.ReadByte)
                                Next

                                bdtoffset = BDTStream.Position
                                If bdtoffset Mod &H10 > 0 Then
                                    padding = &H10 - (bdtoffset Mod &H10)
                                Else
                                    padding = 0
                                End If
                                bdtoffset += padding

                                BDTStream.Position = bdtoffset

                                'fStream.Close()
                                fStream.Dispose()

                            Next

                            'BDTStream.Close()
                            BDTStream.Dispose()

                            output(TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine)

                            'txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

                        Case "BHF4"

                            Dim type As Byte
                            Dim unicode As Byte
                            Dim extendedHeader As Byte
                            ReDim bytes(&H3F)
                            BinderID = fileList(0).Split(",")(1)
                            StrToBytes(fileList(0).Substring(0, 4), 0)
                            StrToBytes(BinderID, &H18)

                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            unicode = flags And &HFF
                            type = (flags And &HFF00) >> 8
                            extendedHeader = flags >> 16
                            For i = 2 To fileList.Length - 1
                                namesEndLoc += EncodeFileNameBND4(fileList(i)).Length - InStr(fileList(i), ",") * 2 + 2
                            Next

                            Select Case type
                                Case &H74, &H54
                                    currFileNameOffset = &H40 + &H24 * numFiles
                                    namesEndLoc += &H40 + &H24 * numFiles
                                    bigEndian = False
                            End Select


                            UIntToBytes(&H10000, &H8)
                            UIntToBytes(&H40, &H10)
                            UIntToBytes(&H24, &H20)
                            UIntToBytes(flags, &H30)
                            UIntToBytes(numFiles, &HC)
                            UIntToBytes(namesEndLoc, &H38)


                            Dim groupCount As UInteger


                            For i As UInteger = numFiles \ 7 To 100000
                                Dim noPrime = False
                                For j As UInteger = 2 To i - 1
                                    If i Mod j = 0 Or i = 2 Then
                                        noPrime = True
                                        Exit For
                                    End If
                                Next
                                If noPrime = False And i > 1 Then
                                    groupCount = i
                                    Exit For
                                End If
                            Next

                            Dim hashLists(groupCount) As List(Of pathHash)


                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i) = New List(Of pathHash)
                            Next

                            Dim hashGroups As New List(Of hashGroup)
                            Dim pathHashes As New List(Of pathHash)

                            If extendedHeader = 4 Then
                                For i As UInteger = 0 To numFiles - 1
                                    Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                    Dim pathHash As pathHash = New pathHash()
                                    If internalFileName(0) <> "\" Then
                                        internalFileName = "\" & internalFileName
                                    End If
                                    Dim hash As UInteger = HashFileName(internalFileName.Replace("\", "/"))

                                    pathHash.hash = hash
                                    pathHash.idx = i
                                    Dim group As UInteger = hash Mod groupCount
                                    hashLists(group).Add(pathHash)
                                Next

                                For i As UInteger = 0 To groupCount - 1
                                    hashLists(i).Sort(Function(x, y) x.hash.CompareTo(y.hash))
                                Next


                                Dim count As UInteger = 0
                                For i As UInteger = 0 To groupCount - 1
                                    Dim index As UInteger = count
                                    For Each pathHash As pathHash In hashLists(i)
                                        pathHashes.Add(pathHash)
                                        count += 1
                                    Next
                                    Dim hashGroup As hashGroup
                                    hashGroup.idx = index
                                    hashGroup.length = count - index
                                    hashGroups.Add(hashGroup)
                                Next

                                Dim extendedPadding As UInteger = namesEndLoc Mod 8
                                If extendedPadding = 0 Then
                                Else
                                    namesEndLoc += 8 - extendedPadding
                                End If

                                ReDim Preserve bytes((namesEndLoc - 1) + &H10 + groupCount * 8 + numFiles * 8)

                                UIntToBytes(namesEndLoc, &H38)
                                UIntToBytes(namesEndLoc + &H10 + groupCount * 8, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(0, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(groupCount, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(&H80810, namesEndLoc)
                                namesEndLoc += 4

                                For i As UInteger = 0 To groupCount - 1
                                    UIntToBytes(hashGroups(i).length, namesEndLoc)
                                    namesEndLoc += 4
                                    UIntToBytes(hashGroups(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next

                                For i As UInteger = 0 To numFiles - 1
                                    UIntToBytes(pathHashes(i).hash, namesEndLoc)
                                    namesEndLoc += 4
                                    UIntToBytes(pathHashes(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next


                            End If


                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, bndfile.Length - 3) & "bdt"

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            BDTStream.Position = 0
                            WriteBytes(BDTStream, StrToBytes("BDF4"))
                            BDTStream.Position = &HA
                            BDTStream.WriteByte(1)
                            BDTStream.Position = &H10
                            BDTStream.WriteByte(&H30)
                            BDTStream.Position = &H18
                            WriteBytes(BDTStream, StrToBytes(BinderID))

                            Dim bdtoffset As UInteger = &H30

                            BDTStream.Position = bdtoffset

                            For i = 0 To numFiles - 1

                                Select Case type
                                    Case &H74, &H54
                                        currFileID = fileList(i + 2).Split(",")(0)
                                        currFileName = fileList(i + 2).Split(",")(1)

                                        Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                                        UIntToBytes(&H40, &H40 + i * &H24)
                                        UIntToBytes(&HFFFFFFFF, &H44 + i * &H24)
                                        UInt64ToBytes(fStream.Length, &H48 + i * &H24)
                                        UInt64ToBytes(fStream.Length, &H50 + i * &H24)
                                        UIntToBytes(bdtoffset, &H58 + i * &H24)
                                        UIntToBytes(currFileID, &H5C + i * &H24)
                                        UIntToBytes(currFileNameOffset, &H60 + i * &H24)

                                        EncodeFileNameBND4(currFileName, currFileNameOffset)
                                        currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2

                                        For j = 0 To fStream.Length - 1
                                            BDTStream.WriteByte(fStream.ReadByte)
                                        Next

                                        bdtoffset = BDTStream.Position

                                        If bdtoffset Mod &H10 > 0 Then
                                            padding = &H10 - (bdtoffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        bdtoffset += padding

                                        BDTStream.Position = bdtoffset

                                        fStream.Close()
                                        fStream.Dispose()





                                End Select

                            Next

                            'BDTStream.Close()
                            BDTStream.Dispose()

                            output(TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine)

                            'txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

                        Case "BND3"
                            ReDim bytes(&H1F)
                            StrToBytes(fileList(0), 0)

                            flags = fileList(1)
                            numFiles = fileList.Length - 2


                            For i = 2 To fileList.Length - 1
                                namesEndLoc += EncodeFileName(fileList(i)).Length - InStr(fileList(i), ",") + 1
                            Next

                            Select Case flags
                                Case &H74000000, &H78000000, &H54000000
                                    currFileNameOffset = &H20 + &H18 * numFiles
                                    namesEndLoc += &H20 + &H18 * numFiles
                                Case &H10100
                                    namesEndLoc = &H20 + &HC * numFiles
                                Case &HE010100
                                    currFileNameOffset = &H20 + &H14 * numFiles
                                    namesEndLoc += &H20 + &H14 * numFiles
                                Case &H2E010100
                                    currFileNameOffset = &H20 + &H18 * numFiles
                                    namesEndLoc += &H20 + &H18 * numFiles
                                Case &H7C000000, &H5C000000
                                    currFileNameOffset = &H20 + &H1C * numFiles
                                    namesEndLoc += &H20 + &H1C * numFiles
                            End Select

                            UIntToBytes(flags, &HC)
                            If flags = &H74000000 Or flags = &H78000000 Or flags = &H54000000 Or flags = &H7C000000 Or flags = &H5C000000 Then bigEndian = False

                            UIntToBytes(numFiles, &H10)
                            UIntToBytes(namesEndLoc, &H14)

                            If namesEndLoc Mod &H10 > 0 Then
                                padding = &H10 - (namesEndLoc Mod &H10)
                            Else
                                padding = 0
                            End If

                            ReDim Preserve bytes(namesEndLoc + padding - 1)

                            currFileOffset = namesEndLoc + padding

                            For i As UInteger = 0 To numFiles - 1
                                Select Case flags
                                    Case &H74000000, &H54000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H40, &H20 + i * &H18)
                                        UIntToBytes(tmpbytes.Length, &H24 + i * &H18)
                                        UIntToBytes(currFileOffset, &H28 + i * &H18)
                                        UIntToBytes(currFileID, &H2C + i * &H18)
                                        UIntToBytes(currFileNameOffset, &H30 + i * &H18)
                                        UIntToBytes(tmpbytes.Length, &H34 + i * &H18)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H78000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H40, &H20 + i * &H18)
                                        UIntToBytes(tmpbytes.Length, &H24 + i * &H18)
                                        UInt64ToBytes(currFileOffset, &H28 + i * &H18)
                                        UIntToBytes(currFileID, &H30 + i * &H18)
                                        UIntToBytes(currFileNameOffset, &H34 + i * &H18)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H7C000000, &H5C000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H40, &H20 + i * &H1C)
                                        UIntToBytes(tmpbytes.Length, &H24 + i * &H1C)
                                        UInt64ToBytes(currFileOffset, &H28 + i * &H1C)
                                        UIntToBytes(currFileID, &H30 + i * &H1C)
                                        UIntToBytes(currFileNameOffset, &H34 + i * &H1C)
                                        UIntToBytes(tmpbytes.Length, &H38 + i * &H1C)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H10100
                                        currFileName = fileList(i + 2)
                                        currFileName = filepath & filename & ".extract\" & currFileName
                                        currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                        currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                        tmpbytes = File.ReadAllBytes(currFilePath & currFileName)
                                        currFileSize = tmpbytes.Length

                                        If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                                            padding = &H10 - (currFileSize Mod &H10)
                                        Else
                                            padding = 0
                                        End If

                                        UIntToBytes(&H2000000, &H20 + i * &HC)
                                        UIntToBytes(currFileSize, &H24 + i * &HC)
                                        UIntToBytes(currFileOffset, &H28 + i * &HC)

                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length + padding

                                    Case &HE010100
                                        currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H2000000, &H20 + i * &H14)
                                        UIntToBytes(tmpbytes.Length, &H24 + i * &H14)
                                        UIntToBytes(currFileOffset, &H28 + i * &H14)
                                        UIntToBytes(currFileID, &H2C + i * &H14)
                                        UIntToBytes(currFileNameOffset, &H30 + i * &H14)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H2E010100
                                        currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H2000000, &H20 + i * &H18)
                                        UIntToBytes(tmpbytes.Length, &H24 + i * &H18)
                                        UIntToBytes(currFileOffset, &H28 + i * &H18)
                                        UIntToBytes(currFileID, &H2C + i * &H18)
                                        UIntToBytes(currFileNameOffset, &H30 + i * &H18)
                                        UIntToBytes(tmpbytes.Length, &H34 + i * &H18)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                End Select
                            Next
                        Case "BND4"

                            'Reversing and hash grouping code by TKGP
                            'https://github.com/JKAnderson/SoulsFormats

                            Dim type As Byte
                            Dim unicode As Byte
                            Dim extendedHeader As Byte
                            Dim entryLength As UInteger = &H24
                            ReDim bytes(&H3F)
                            StrToBytes(fileList(0).Substring(0, 4), 0)
                            StrToBytes(fileList(0).Substring(4), &H18)
                            If IsRegulation Then
                                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"
                            End If

                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            unicode = flags And &HFF
                            type = (flags And &HFF00) >> 8
                            extendedHeader = flags >> 16
                            For i = 2 To fileList.Length - 1
                                If unicode > 0 Then
                                    namesEndLoc += EncodeFileNameBND4(fileList(i)).Length - InStr(fileList(i), ",") * 2 + 2
                                Else
                                    namesEndLoc += EncodeFileName(fileList(i)).Length - InStr(fileList(i), ",") + 1
                                End If

                            Next

                            Select Case type
                                Case &H74, &H54
                                    entryLength = &H24
                                    currFileNameOffset = &H40 + entryLength * numFiles
                                    namesEndLoc += &H40 + entryLength * numFiles
                                    bigEndian = False

                                Case &H20
                                    entryLength = &H20
                                    currFileNameOffset = &H40 + entryLength * numFiles
                                    namesEndLoc += &H40 + entryLength * numFiles
                                    bigEndian = False
                            End Select


                            UIntToBytes(&H10000, &H8)
                            UIntToBytes(&H40, &H10)
                            UIntToBytes(entryLength, &H20)
                            UIntToBytes(flags, &H30)
                            UIntToBytes(numFiles, &HC)


                            Dim groupCount As UInteger


                            For i As UInteger = numFiles \ 7 To 100000
                                Dim noPrime = False
                                For j As UInteger = 2 To i - 1
                                    If i Mod j = 0 Or i = 2 Then
                                        noPrime = True
                                        Exit For
                                    End If
                                Next
                                If noPrime = False And i > 1 Then
                                    groupCount = i
                                    Exit For
                                End If
                            Next

                            Dim hashLists(groupCount) As List(Of pathHash)


                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i) = New List(Of pathHash)
                            Next

                            Dim hashGroups As New List(Of hashGroup)
                            Dim pathHashes As New List(Of pathHash)

                            If extendedHeader = 4 Then
                                For i As UInteger = 0 To numFiles - 1
                                    Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                    Dim pathHash As pathHash = New pathHash()
                                    If internalFileName(0) <> "\" Then
                                        internalFileName = "\" & internalFileName
                                    End If
                                    Dim hash As UInteger = HashFileName(internalFileName.Replace("\", "/"))

                                    pathHash.hash = hash
                                    pathHash.idx = i
                                    Dim group As UInteger = hash Mod groupCount
                                    hashLists(group).Add(pathHash)
                                Next

                                For i As UInteger = 0 To groupCount - 1
                                    hashLists(i).Sort(Function(x, y) x.hash.CompareTo(y.hash))
                                Next


                                Dim count As UInteger = 0
                                For i As UInteger = 0 To groupCount - 1
                                    Dim index As UInteger = count
                                    For Each pathHash As pathHash In hashLists(i)
                                        pathHashes.Add(pathHash)
                                        count += 1
                                    Next
                                    Dim hashGroup As hashGroup
                                    hashGroup.idx = index
                                    hashGroup.length = count - index
                                    hashGroups.Add(hashGroup)
                                Next

                                Dim extendedPadding As UInteger = namesEndLoc Mod 8
                                If extendedPadding = 0 Then
                                Else
                                    namesEndLoc += 8 - extendedPadding
                                End If


                                ReDim Preserve bytes((namesEndLoc - 1) + &H10 + groupCount * 8 + numFiles * 8)

                                UIntToBytes(namesEndLoc, &H38)
                                UIntToBytes(namesEndLoc + &H10 + groupCount * 8, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(0, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(groupCount, namesEndLoc)
                                namesEndLoc += 4
                                UIntToBytes(&H80810, namesEndLoc)
                                namesEndLoc += 4

                                For i As UInteger = 0 To groupCount - 1
                                    UIntToBytes(hashGroups(i).length, namesEndLoc)
                                    namesEndLoc += 4
                                    UIntToBytes(hashGroups(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next

                                For i As UInteger = 0 To numFiles - 1
                                    UIntToBytes(pathHashes(i).hash, namesEndLoc)
                                    namesEndLoc += 4
                                    UIntToBytes(pathHashes(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next


                            End If

                            If namesEndLoc Mod &H10 > 0 Then
                                padding = &H10 - (namesEndLoc Mod &H10)
                            Else
                                padding = 0
                            End If

                            ReDim Preserve bytes(namesEndLoc + padding - 1)

                            currFileOffset = namesEndLoc + padding

                            If type = &H20 Then
                                namesEndLoc += padding
                            End If

                            UIntToBytes(namesEndLoc, &H28)


                            For i As UInteger = 0 To numFiles - 1
                                Select Case type
                                    Case &H74, &H54
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H40, &H40 + i * entryLength)
                                        UIntToBytes(&HFFFFFFFF, &H44 + i * entryLength)
                                        UInt64ToBytes(tmpbytes.Length, &H48 + i * entryLength)
                                        'UIntToBytes(0, &H4C + i * &H24)
                                        UInt64ToBytes(tmpbytes.Length, &H50 + i * entryLength)
                                        'UIntToBytes(0, &H54 + i * &H24)
                                        UIntToBytes(currFileOffset, &H58 + i * entryLength)
                                        UIntToBytes(currFileID, &H5C + i * entryLength)
                                        UIntToBytes(currFileNameOffset, &H60 + i * entryLength)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))

                                        If unicode > 0 Then
                                            EncodeFileNameBND4(internalFileName, currFileNameOffset)
                                            currFileNameOffset += EncodeFileNameBND4(internalFileName).Length + 2
                                        Else
                                            EncodeFileName(internalFileName, currFileNameOffset)
                                            currFileNameOffset += EncodeFileName(internalFileName).Length + 1
                                        End If


                                    Case &H20
                                        'Save files
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        UIntToBytes(&H50, &H40 + i * entryLength)
                                        UIntToBytes(&HFFFFFFFF, &H44 + i * entryLength)
                                        UInt64ToBytes(tmpbytes.Length, &H48 + i * entryLength)
                                        UIntToBytes(currFileOffset, &H50 + i * entryLength)
                                        UIntToBytes(currFileNameOffset, &H54 + i * entryLength)

                                        If tmpbytes.Length Mod &H10 > 0 Then
                                            padding = &H10 - (tmpbytes.Length Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        If i = numFiles - 1 Then padding = 0
                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length
                                        If currFileOffset Mod &H10 > 0 Then
                                            padding = &H10 - (currFileOffset Mod &H10)
                                        Else
                                            padding = 0
                                        End If
                                        currFileOffset += padding

                                        Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))

                                        If unicode > 0 Then
                                            EncodeFileNameBND4(internalFileName, currFileNameOffset)
                                            currFileNameOffset += EncodeFileNameBND4(internalFileName).Length + 2
                                        Else
                                            EncodeFileName(internalFileName, currFileNameOffset)
                                            currFileNameOffset += EncodeFileName(internalFileName).Length + 1
                                        End If

                                End Select
                            Next


                        Case "TPF"
                            'TODO:  Handle m10_9999 (PC) format
                            Dim currFileFlags1
                            Dim currFileFlags2
                            Dim totalFileSize = 0
                            ReDim bytes(&HF)
                            StrToBytes(fileList(0), 0)

                            flags = fileList(1)

                            If flags = &H2010200 Or flags = &H201000 Or flags = &H2020200 Or flags = &H2030200 Then
                                ' Demon's Souls/Dark Souls (headerless DDS)

                                Dim currFileFormat As UInteger = 0
                                Dim currFileIsCubemap As Boolean = False
                                Dim currFileMipMaps As UInteger = 0
                                Dim currFileHeight As UShort = 0
                                Dim currFileWidth As UShort = 0
                                Dim cubeBytes() As Byte

                                bigEndian = True

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H20

                                For i = 2 To fileList.Length - 1
                                    namesEndLoc += EncodeFileName(fileList(i)).Length - InStrRev(fileList(i), ",") - 3
                                Next

                                UIntToBytes(numFiles, &H8)
                                UIntToBytes(flags, &HC)

                                If namesEndLoc Mod &H10 > 0 Then
                                    padding = &H10 - (namesEndLoc Mod &H10)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                UIntToBytes(currFileOffset, &H10)

                                currFileNameOffset = &H10 + &H20 * numFiles

                                For i = 0 To numFiles - 1

                                    Dim FileInfo As String() = Split(fileList(i + 2), ",")

                                    currFileFormat = Convert.ToUInt32(FileInfo(0))
                                    currFileIsCubemap = Convert.ToUInt32(FileInfo(1))
                                    currFileMipMaps = Convert.ToUInt32(FileInfo(2))
                                    currFileFlags1 = Convert.ToUInt32(FileInfo(3))
                                    currFileFlags2 = Convert.ToUInt32(FileInfo(4))
                                    currFileName = filepath & filename & ".extract\" & FileInfo(5)

                                    tmpbytes = File.ReadAllBytes(currFileName)

                                    currFileSize = tmpbytes.Length - &H80

                                    currFileWidth = BitConverter.ToUInt16(tmpbytes, 12)
                                    currFileHeight = BitConverter.ToUInt16(tmpbytes, 16)

                                    If (currFileFormat = 10 Or currFileFormat = 9) Then
                                        If currFileIsCubemap Then
                                            Dim linearSize = BitConverter.ToUInt32(tmpbytes, 20)
                                            ReDim outBytes(linearSize * (4 / 3) - 1)
                                            ReDim cubeBytes(currFileSize * (4 / 3) - 1)
                                            'Swizzle cube faces separately
                                            For j As UInteger = 0 To 5
                                                ReDim inBytes(linearSize - 1)
                                                Array.Copy(tmpbytes, &H80 + j * linearSize, inBytes, 0, linearSize)
                                                ExtendInBytes(True)
                                                SwizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                                writeOffset = 0
                                                Array.Copy(outBytes, 0, cubeBytes, CType(j * linearSize * (4 / 3), UInteger), CType(linearSize * (4 / 3), UInteger))
                                            Next
                                            currFileSize = cubeBytes.Length
                                        Else
                                            ReDim inBytes(currFileSize - 1)
                                            ReDim outBytes(currFileSize * (4 / 3) - 1)
                                            ddsWidth = currFileWidth
                                            Array.Copy(tmpbytes, &H80, inBytes, 0, currFileSize)
                                            ExtendInBytes(False)
                                            SwizzleDDSBytes(currFileWidth, currFileHeight, 0, 2)
                                            currFileSize = outBytes.Length
                                            writeOffset = 0
                                        End If
                                    ElseIf currFileIsCubemap Then
                                        tmpbytes = tmpbytes.Skip(&H80).ToArray()
                                        AddMemeBytes(tmpbytes)
                                        currFileSize = tmpbytes.Length
                                    Else
                                        tmpbytes = tmpbytes.Skip(&H80).ToArray()
                                    End If

                                    If currFileSize Mod &H10 > 0 Then
                                        padding = &H10 - (currFileSize Mod &H10)
                                    Else
                                        padding = 0
                                    End If

                                    UIntToBytes(currFileOffset, &H10 + i * &H20)
                                    UIntToBytes(currFileSize, &H14 + i * &H20)
                                    bytes(&H18 + i * &H20) = currFileFormat
                                    bytes(&H19 + i * &H20) = Convert.ToByte(currFileIsCubemap)
                                    bytes(&H1A + i * &H20) = currFileMipMaps
                                    UInt16ToBytes(currFileWidth, &H1C + i * &H20)
                                    UInt16ToBytes(currFileHeight, &H1E + i * &H20)
                                    UIntToBytes(currFileFlags1, &H20 + i * &H20)
                                    UIntToBytes(currFileFlags2, &H24 + i * &H20)
                                    UIntToBytes(currFileNameOffset, &H28 + i * &H20)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    If currFileFormat = 10 Or currFileFormat = 9 Then
                                        If currFileIsCubemap Then
                                            InsBytes(cubeBytes, currFileOffset)
                                        Else
                                            InsBytes(outBytes, currFileOffset)
                                        End If
                                    Else
                                        InsBytes(tmpbytes, currFileOffset)
                                    End If

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = FileInfo(5).Substring(0, FileInfo(5).Length - ".dds".Length)
                                    EncodeFileName(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileName(currFileName).Length + 1
                                Next

                                UIntToBytes(totalFileSize, &H4)
                            ElseIf flags = &H20300 Or flags = &H20304 Then
                                ' Dark Souls

                                bigEndian = False

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H14

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileName(currFileName).Length + 1
                                Next

                                UIntToBytes(numFiles, &H8)
                                UIntToBytes(flags, &HC)

                                If namesEndLoc Mod &H10 > 0 Then
                                    padding = &H10 - (namesEndLoc Mod &H10)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                currFileNameOffset = &H10 + &H14 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = fileList(i + 2)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFilePath = filepath & filename & ".extract\"
                                    currFileName = currFileName

                                    tmpbytes = File.ReadAllBytes(currFilePath & currFileName)

                                    currFileSize = tmpbytes.Length
                                    If currFileSize Mod &H10 > 0 Then
                                        padding = &H10 - (currFileSize Mod &H10)
                                    Else
                                        padding = 0
                                    End If

                                    Dim words() As String = fileList(i + 2).Split(",")
                                    currFileFlags1 = words(0)
                                    currFileFlags2 = words(1)

                                    UIntToBytes(currFileOffset, &H10 + i * &H14)
                                    UIntToBytes(currFileSize, &H14 + i * &H14)
                                    UIntToBytes(currFileFlags1, &H18 + i * &H14)
                                    UIntToBytes(currFileNameOffset, &H1C + i * &H14)
                                    UIntToBytes(currFileFlags2, &H20 + i * &H14)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileName(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileName(currFileName).Length + 1
                                Next

                                UIntToBytes(totalFileSize, &H4)
                            ElseIf flags = &H10300 Then
                                ' Dark Souls III

                                Dim currFileFormat As UInteger = 0
                                Dim currFileIsCubemap As Boolean = False
                                Dim currFileMipMaps As UInteger = 0

                                bigEndian = False

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H14

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileNameBND4(currFileName).Length + 2
                                Next

                                UIntToBytes(numFiles, &H8)
                                UIntToBytes(flags, &HC)

                                'If namesEndLoc Mod &H10 > 0 Then
                                'padding = &H10 - (namesEndLoc Mod &H10)
                                'Else
                                'padding = 0
                                'End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                currFileNameOffset = &H10 + &H14 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = fileList(i + 2)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFilePath = filepath & filename & ".extract\"
                                    currFileName = currFileName

                                    tmpbytes = File.ReadAllBytes(currFilePath & currFileName)

                                    currFileSize = tmpbytes.Length
                                    'If currFileSize Mod &H10 > 0 Then
                                    'padding = &H10 - (currFileSize Mod &H10)
                                    'Else
                                    'padding = 0
                                    'End If

                                    Dim words() As String = fileList(i + 2).Split(",")
                                    currFileFormat = words(0)
                                    currFileIsCubemap = words(1)
                                    currFileMipMaps = words(2)
                                    currFileFlags1 = words(3)

                                    UIntToBytes(currFileOffset, &H10 + i * &H14)
                                    UIntToBytes(currFileSize, &H14 + i * &H14)
                                    bytes(&H18 + i * &H14) = currFileFormat
                                    bytes(&H19 + i * &H14) = Convert.ToByte(currFileIsCubemap)
                                    bytes(&H1A + i * &H14) = currFileMipMaps
                                    UIntToBytes(currFileNameOffset, &H1C + i * &H14)
                                    UIntToBytes(currFileFlags1, &H20 + i * &H14)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileNameBND4(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2
                                Next

                                UIntToBytes(totalFileSize, &H4)
                            ElseIf flags = &H10304 Then
                                ' Dark Souls III (headerless DDS)

                                Dim currFileFormat As UInteger = 0
                                Dim currFileIsCubemap As Boolean = False
                                Dim currFileMipMaps As UInteger = 0
                                Dim currFileArraySize As UInteger = 0
                                Dim currFileDxgiFormat As UInteger = 0
                                Dim currFileWidth As UShort = 0
                                Dim currFileHeight As UShort = 0
                                Dim paddedWidth As UInteger = 0
                                Dim paddedHeight As UInteger = 0
                                Dim paddedSize As UInteger = 0
                                Dim Width As UInteger
                                Dim Height As UInteger
                                Dim BlockSize As UInteger = 0
                                Dim copyOffset As UInteger = 0

                                bigEndian = False

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H24

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileNameBND4(currFileName).Length + 2
                                Next

                                UIntToBytes(numFiles, &H8)
                                UIntToBytes(flags, &HC)

                                If namesEndLoc Mod &H10 > 0 Then
                                    padding = &H10 - (namesEndLoc Mod &H10)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                currFileNameOffset = &H10 + &H24 * numFiles

                                For i = 0 To numFiles - 1

                                    Dim FileInfo As String() = Split(fileList(i + 2), ",")
                                    Dim currMipMapOffset As UInteger = 0

                                    currFileSize = 0
                                    currFileFormat = Convert.ToUInt32(FileInfo(0))
                                    currFileIsCubemap = Convert.ToUInt32(FileInfo(1))
                                    currFileMipMaps = Convert.ToUInt32(FileInfo(2))
                                    currFileArraySize = Convert.ToUInt32(FileInfo(3))
                                    currFileDxgiFormat = Convert.ToUInt32(FileInfo(4))
                                    currFileName = filepath & filename & ".extract\" & FileInfo(5)

                                    tmpbytes = File.ReadAllBytes(currFilePath & currFileName)

                                    currFileHeight = BitConverter.ToUInt16(tmpbytes, 12)
                                    currFileWidth = BitConverter.ToUInt16(tmpbytes, 16)

                                    Width = currFileWidth
                                    Height = currFileHeight

                                    If currFileFormat = 22 Or currFileFormat = 105 Then
                                        copyOffset = &H80
                                    Else
                                        copyOffset = &H94
                                    End If

                                    Select Case (currFileFormat)
                                        Case 105
                                            BlockSize = 4
                                        Case 22, 25
                                            BlockSize = 8
                                        Case 0, 1, 103, 108, 109
                                            BlockSize = 8
                                        Case 5, 100, 102, 106, 107, 110
                                            BlockSize = 16
                                        Case Else
                                            MsgBox("Format " & currFileFormat & " in file " & currFileName & " not implemented. Aborting :)")
                                            SyncLock workLock
                                                work = False
                                            End SyncLock
                                            Return
                                    End Select

                                    If currFileIsCubemap Then
                                        Dim ArrayEntrySize = 0
                                        For j As UInteger = 0 To currFileMipMaps - 1
                                            If currFileFormat = 22 Then
                                                ArrayEntrySize += Width * Height * BlockSize
                                            Else
                                                ArrayEntrySize += Math.Ceiling(Width / 4) * Math.Ceiling(Height / 4) * BlockSize
                                            End If

                                            If Width > 1 Then
                                                Width /= 2
                                            End If
                                            If Height > 1 Then
                                                Height /= 2
                                            End If
                                        Next

                                        Width = currFileWidth
                                        Height = currFileHeight

                                        For j As UInteger = 0 To currFileMipMaps - 1
                                            If currFileFormat = 22 Then
                                                paddedWidth = Math.Ceiling(Width / 8) * 8
                                                paddedHeight = Math.Ceiling(Height / 8) * 8
                                                paddedSize = paddedWidth * paddedHeight * BlockSize
                                            Else
                                                paddedWidth = Math.Ceiling(Width / 32) * 32
                                                paddedHeight = Math.Ceiling(Height / 32) * 32
                                                paddedSize = Math.Ceiling(paddedWidth / 4) * Math.Ceiling(paddedHeight / 4) * BlockSize
                                            End If

                                            ddsWidth = paddedWidth
                                            ReDim outBytes(paddedSize - 1)
                                            ReDim inBytes(paddedSize - 1)
                                            ReDim Preserve bytes(bytes.Length + paddedSize * currFileArraySize + 2 * paddedSize - 1)

                                            For k As UInteger = 0 To currFileArraySize - 1
                                                Dim tempOffset = copyOffset

                                                If currFileFormat = 22 Then
                                                    For l As UInteger = 0 To Height - 1
                                                        Array.Copy(tmpbytes, copyOffset, inBytes, l * paddedWidth * BlockSize, Width * BlockSize)
                                                        copyOffset += Width * BlockSize
                                                    Next
                                                Else
                                                    For l As UInteger = 0 To Math.Ceiling(Height / 4) - 1
                                                        Array.Copy(tmpbytes, copyOffset, inBytes, CType(l * Math.Ceiling(paddedWidth / 4) * BlockSize, UInteger), CType(Math.Ceiling(Width / 4) * BlockSize, UInteger))
                                                        copyOffset += Math.Ceiling(Width / 4) * BlockSize
                                                    Next
                                                End If

                                                copyOffset = tempOffset

                                                SwizzleDDSBytesPS4(paddedWidth, paddedHeight, currFileFormat)

                                                InsBytes(outBytes, currFileOffset + currMipMapOffset)
                                                currMipMapOffset += paddedSize
                                                copyOffset += ArrayEntrySize
                                            Next

                                            currMipMapOffset += paddedSize * 2
                                            copyOffset -= ArrayEntrySize * currFileArraySize

                                            If currFileFormat = 22 Then
                                                copyOffset += Width * Height * BlockSize
                                            Else
                                                copyOffset += Math.Ceiling(Width / 4) * Math.Ceiling(Height / 4) * BlockSize
                                            End If

                                            If Width > 1 Then
                                                Width /= 2
                                            End If
                                            If Height > 1 Then
                                                Height /= 2
                                            End If
                                        Next
                                    Else
                                        For j As UInteger = 0 To currFileMipMaps - 1
                                            If currFileFormat = 105 Then
                                                paddedWidth = Width
                                                paddedHeight = Height
                                                paddedSize = paddedWidth * paddedHeight * BlockSize
                                            Else
                                                paddedWidth = Math.Ceiling(Width / 32) * 32
                                                paddedHeight = Math.Ceiling(Height / 32) * 32
                                                paddedSize = Math.Ceiling(paddedWidth / 4) * Math.Ceiling(paddedHeight / 4) * BlockSize
                                            End If

                                            ddsWidth = paddedWidth
                                            ReDim outBytes(paddedSize - 1)
                                            ReDim inBytes(paddedSize - 1)

                                            If currFileFormat = 105 Then
                                                Array.Copy(tmpbytes, copyOffset, inBytes, 0, paddedSize)
                                            Else
                                                For k As UInteger = 0 To Math.Ceiling(Height / 4) - 1
                                                    Array.Copy(tmpbytes, copyOffset, inBytes, CType(k * Math.Ceiling(paddedWidth / 4) * BlockSize, UInteger), CType(Math.Ceiling(Width / 4) * BlockSize, UInteger))
                                                    copyOffset += Math.Ceiling(Width / 4) * BlockSize
                                                Next
                                            End If

                                            SwizzleDDSBytesPS4(Width, Height, currFileFormat)

                                            ReDim Preserve bytes(bytes.Length + paddedSize - 1)

                                            InsBytes(outBytes, currFileOffset + currMipMapOffset)
                                            currMipMapOffset += paddedSize
                                            If Width > 1 Then
                                                Width /= 2
                                            End If
                                            If Height > 1 Then
                                                Height /= 2
                                            End If
                                        Next
                                    End If

                                    currFileSize = currMipMapOffset
                                    UIntToBytes(currFileOffset, &H10 + i * &H24)
                                    UIntToBytes(currFileSize, &H14 + i * &H24)
                                    bytes(&H18 + i * &H24) = currFileFormat
                                    bytes(&H19 + i * &H24) = Convert.ToByte(currFileIsCubemap)
                                    bytes(&H1A + i * &H24) = currFileMipMaps
                                    UInt16ToBytes(currFileWidth, &H1C + i * &H24)
                                    UInt16ToBytes(currFileHeight, &H1E + i * &H24)
                                    UIntToBytes(currFileArraySize, &H20 + i * &H24)
                                    UIntToBytes(&HD, &H24 + i * &H24)
                                    UIntToBytes(currFileNameOffset, &H28 + i * &H24)

                                    currFileName = FileInfo(5).Substring(0, FileInfo(5).Length - ".dds".Length)
                                    EncodeFileNameBND4(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2

                                    UIntToBytes(currFileDxgiFormat, &H30 + i * &H24)

                                    currFileOffset += currMipMapOffset
                                Next

                                UIntToBytes(currFileOffset - namesEndLoc - padding, &H4)

                            End If
                    End Select

                    If Not IsRegulation Then
                        File.WriteAllBytes(filepath & filename, bytes)
                        output(TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine)
                    End If


                End If


                If DCX Or IsRegulation Then
                    bigEndian = True
                    If IsRegulation Then
                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                    Else
                        filename = filename & ".dcx"

                        If Not File.Exists(filepath & filename & ".bak") Then
                            File.WriteAllBytes(filepath & filename & ".bak", dcxBytes)
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                            output(TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                            output(TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine)
                        End If
                    End If

                    fileList = File.ReadAllLines(filepath & filename & ".info.txt")

                    Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
                        Case "EDGE"
                            Dim chunkBytes(&H10000) As Byte
                            Dim cmpChunkBytes() As Byte
                            Dim zipBytes() As Byte = {}

                            If fileList.Length > 2 Then
                                currFileName = filepath + fileList(2)
                            Else
                                currFileName = filepath + fileList(1)
                            End If
                            tmpbytes = File.ReadAllBytes(currFileName)

                            currFileSize = tmpbytes.Length

                            ReDim bytes(&H83)

                            Dim fileRemaining As Integer = tmpbytes.Length
                            Dim fileDone As Integer = 0
                            Dim fileToDo As Integer = 0
                            Dim chunks = 0
                            Dim lastchunk = 0

                            While fileRemaining > 0
                                chunks += 1

                                If fileRemaining > &H10000 Then
                                    fileToDo = &H10000
                                Else
                                    fileToDo = fileRemaining
                                End If


                                Array.Copy(tmpbytes, fileDone, chunkBytes, 0, fileToDo)
                                cmpChunkBytes = Compress(chunkBytes, True)

                                lastchunk = zipBytes.Length

                                If lastchunk Mod &H10 > 0 Then
                                    padding = &H10 - (lastchunk Mod &H10)
                                Else
                                    padding = 0
                                End If
                                lastchunk += padding

                                ReDim Preserve zipBytes(lastchunk + cmpChunkBytes.Length)
                                Array.Copy(cmpChunkBytes, 0, zipBytes, lastchunk, cmpChunkBytes.Length)


                                fileDone += fileToDo
                                fileRemaining -= fileToDo

                                ReDim Preserve bytes(bytes.Length + &H10)

                                UIntToBytes(lastchunk, &H64 + chunks * &H10)
                                UIntToBytes(cmpChunkBytes.Length, &H68 + chunks * &H10)
                                UIntToBytes(&H1, &H6C + chunks * &H10)

                            End While
                            ReDim Preserve bytes(bytes.Length + zipBytes.Length)

                            StrToBytes("DCX", &H0)
                            UIntToBytes(&H10000, &H4)
                            UIntToBytes(&H18, &H8)
                            UIntToBytes(&H24, &HC)
                            UIntToBytes(&H24, &H10)
                            UIntToBytes(&H50 + chunks * &H10, &H14)
                            StrToBytes("DCS", &H18)
                            UIntToBytes(currFileSize, &H1C)
                            UIntToBytes(bytes.Length - (&H70 + chunks * &H10), &H20)
                            StrToBytes("DCP", &H24)
                            StrToBytes("EDGE", &H28)
                            UIntToBytes(&H20, &H2C)
                            UIntToBytes(&H9000000, &H30)
                            UIntToBytes(&H10000, &H34)

                            UIntToBytes(&H100100, &H40)
                            StrToBytes("DCA", &H44)
                            UIntToBytes(chunks * &H10 + &H2C, &H48)
                            StrToBytes("EgdT", &H4C)
                            UIntToBytes(&H10100, &H50)
                            UIntToBytes(&H24, &H54)
                            UIntToBytes(&H10, &H58)
                            UIntToBytes(&H10000, &H5C)
                            UIntToBytes(tmpbytes.Length Mod &H10000, &H60)
                            UIntToBytes(&H24 + chunks * &H10, &H64)
                            UIntToBytes(chunks, &H68)
                            UIntToBytes(&H100000, &H6C)

                            Array.Copy(zipBytes, 0, bytes, &H70 + chunks * &H10, zipBytes.Length)
                        Case "DFLT"
                            Dim cmpBytes() As Byte
                            Dim zipBytes() As Byte = {}

                            If fileList.Length > 2 Then
                                currFileName = filepath + fileList(2)
                            Else
                                currFileName = filepath + fileList(1)
                            End If
                            If IsRegulation Then
                                tmpbytes = bytes
                            Else
                                tmpbytes = File.ReadAllBytes(currFileName)
                            End If

                            currFileSize = tmpbytes.Length

                            ReDim bytes(&H4C)


                            cmpBytes = Compress(tmpbytes, False)

                            ReDim Preserve bytes(bytes.Length + cmpBytes.Length)

                            StrToBytes("DCX", &H0)
                            UIntToBytes(&H10000, &H4)
                            UIntToBytes(&H18, &H8)
                            UIntToBytes(&H24, &HC)
                            'UIntToBytes(&H24, &H10)
                            'UIntToBytes(&H2C, &H14)
                            If fileList.Length > 2 Then
                                UInt64ToBytes(fileList(1), &H10)
                            Else
                                UIntToBytes(&H24, &H10)
                                UIntToBytes(&H2C, &H14)
                            End If
                            StrToBytes("DCS", &H18)
                            UIntToBytes(currFileSize, &H1C)
                            UIntToBytes(cmpBytes.Length + 2, &H20)
                            StrToBytes("DCP", &H24)
                            StrToBytes("DFLT", &H28)
                            UIntToBytes(&H20, &H2C)
                            UIntToBytes(&H9000000, &H30)

                            UIntToBytes(&H10100, &H40)
                            StrToBytes("DCA", &H44)
                            UIntToBytes(&H8, &H48)
                            UIntToBytes(&H78DA0000, &H4C)


                            Array.Copy(cmpBytes, 0, bytes, &H4E, cmpBytes.Length)
                    End Select

                    If IsRegulation Then
                        output(TimeOfDay & " - Beginning encryption of regulation file." & Environment.NewLine)
                        bytes = EncryptRegulationFile(bytes)
                        output(TimeOfDay & " - Finished encryption of regulation file." & Environment.NewLine)
                    End If

                    File.WriteAllBytes(filepath & filename, bytes)

                    output(TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine)

                    'txtInfo.Text += TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine
                End If


            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            output(TimeOfDay & " - Unhandled exception - " & ex.Message & ex.StackTrace & Environment.NewLine)
        End Try

        SyncLock workLock
            work = False
        End SyncLock

    End Sub

    Public Function Decompress(ByVal cmpBytes() As Byte) As Byte()
        Dim sourceFile As MemoryStream = New MemoryStream(cmpBytes)
        Dim destFile As MemoryStream = New MemoryStream()
        Dim compStream As New DeflateStream(sourceFile, CompressionMode.Decompress)
        Dim myByte As Integer = compStream.ReadByte()

        While myByte <> -1
            destFile.WriteByte(CType(myByte, Byte))
            myByte = compStream.ReadByte()
        End While

        destFile.Close()
        sourceFile.Close()

        Return destFile.ToArray
    End Function

    Private Function Adler32(ByRef input) As UInteger
        Dim a As UInteger = 1
        Dim b As UInteger = 0

        For Each elem As Byte In input
            a = (a + elem) Mod 65521
            b = (b + a) Mod 65521
        Next

        Return (b << 16) Or a
    End Function

    Public Function Compress(ByVal cmpBytes() As Byte, ByVal IsEDGE As Boolean) As Byte()
        Dim ms As New MemoryStream()
        Dim zipStream As Stream = Nothing

        zipStream = New DeflateStream(ms, CompressionMode.Compress, True)
        zipStream.Write(cmpBytes, 0, cmpBytes.Length)
        zipStream.Close()

        ms.Position = 0

        Dim outBytes(ms.Length + 3) As Byte

        If Not IsEDGE Then
            Dim adlerBytes As Byte() = BitConverter.GetBytes(Adler32(cmpBytes))
            Array.Reverse(adlerBytes)
            Array.Copy(adlerBytes, 0, outBytes, ms.Length, 4)
        End If

        ms.Read(outBytes, 0, ms.Length)


        Return outBytes
    End Function

    Private Function MakeDDSHeader(format As UInteger, IsCubemap As Boolean, mipMapCount As UInteger, height As UInteger, width As UInteger) As DDS_HEADER
        Dim ddsHeader As New DDS_HEADER
        ReDim ddsHeader.dwReserved1(10)
        ddsHeader.dwSize = &H7C
        ddsHeader.dwFlags = &HA1007
        ddsHeader.dwHeight = height
        ddsHeader.dwWidth = width

        Select Case format
            Case 0, 1
                'DXT1, DXT1A
                ddsHeader.ddspf.dwFourCC = &H31545844
                ddsHeader.dwPitchOrLinearSize = (height / 4) * (width / 4) * 8
                ddsHeader.ddspf.dwFlags = 4
                ddsHeader.ddspf.dwRGBBitCount = 0
                ddsHeader.ddspf.dwRBitMask = 0
                ddsHeader.ddspf.dwGBitMask = 0
                ddsHeader.ddspf.dwBBitMask = 0
            Case 5
                'DXT5
                ddsHeader.ddspf.dwFourCC = &H35545844
                ddsHeader.dwPitchOrLinearSize = (height / 4) * (width / 4) * 16
                ddsHeader.ddspf.dwFlags = 4
                ddsHeader.ddspf.dwRGBBitCount = 0
                ddsHeader.ddspf.dwRBitMask = 0
                ddsHeader.ddspf.dwGBitMask = 0
                ddsHeader.ddspf.dwBBitMask = 0
            Case 9, 10
                'Uncompressed?
                'This needs deswizzling
                ddsHeader.ddspf.dwFourCC = 0
                ddsHeader.dwPitchOrLinearSize = (height * width) * 3
                ddsHeader.ddspf.dwFlags = &H40
                ddsHeader.ddspf.dwRGBBitCount = &H18
                ddsHeader.ddspf.dwRBitMask = &HFF0000
                ddsHeader.ddspf.dwGBitMask = &HFF00
                ddsHeader.ddspf.dwBBitMask = &HFF
            Case Else
                '???
        End Select

        If IsCubemap Then
            ddsHeader.dwCaps2 = &HFE00
        Else
            ddsHeader.dwCaps2 = 0
        End If

        If mipMapCount Then
            ddsHeader.dwMipMapCount = mipMapCount
        Else
            ddsHeader.dwMipMapCount = 1 + Math.Floor(Math.Log(Math.Max(ddsHeader.dwWidth, ddsHeader.dwHeight), 2))
        End If

        ddsHeader.ddspf.dwSize = &H20

        ddsHeader.dwCaps = &H401008
        Return ddsHeader
    End Function

    Private Function MakeDDSHeaderPS4(format As UInteger, IsCubemap As Boolean, mipMapCount As UInteger, height As UInteger, width As UInteger) As DDS_HEADER
        Dim ddsHeader As New DDS_HEADER
        ReDim ddsHeader.dwReserved1(10)
        ddsHeader.dwSize = &H7C
        ddsHeader.dwFlags = &HA1007
        ddsHeader.dwHeight = height
        ddsHeader.dwWidth = width

        Select Case format
            Case 0, 1, 25, 103, 108, 109
                ddsHeader.ddspf.dwFourCC = &H30315844
                ddsHeader.dwPitchOrLinearSize = (height / 4) * (width / 4) * 8
                ddsHeader.ddspf.dwFlags = 4
                ddsHeader.ddspf.dwRGBBitCount = 0
                ddsHeader.ddspf.dwRBitMask = 0
                ddsHeader.ddspf.dwGBitMask = 0
                ddsHeader.ddspf.dwBBitMask = 0
                ddsHeader.dwCaps = &H401008
            Case 5, 100, 102, 106, 107, 110
                ddsHeader.ddspf.dwFourCC = &H30315844
                ddsHeader.dwPitchOrLinearSize = (height / 4) * (width / 4) * 16
                ddsHeader.ddspf.dwFlags = 4
                ddsHeader.ddspf.dwRGBBitCount = 0
                ddsHeader.ddspf.dwRBitMask = 0
                ddsHeader.ddspf.dwGBitMask = 0
                ddsHeader.ddspf.dwBBitMask = 0
                ddsHeader.dwCaps = &H401008
            Case 22
                ddsHeader.ddspf.dwFourCC = &H71
                ddsHeader.dwPitchOrLinearSize = (height / 4) * (width / 4) * 8
                ddsHeader.ddspf.dwFlags = 4
                ddsHeader.ddspf.dwRGBBitCount = 0
                ddsHeader.ddspf.dwRBitMask = 0
                ddsHeader.ddspf.dwGBitMask = 0
                ddsHeader.ddspf.dwBBitMask = 0
                ddsHeader.dwCaps = &H401008
            Case 105
                ddsHeader.dwFlags = &H100F
                ddsHeader.ddspf.dwFourCC = 0
                ddsHeader.dwPitchOrLinearSize = &H40
                ddsHeader.ddspf.dwFlags = &H41
                ddsHeader.ddspf.dwRGBBitCount = &H20
                ddsHeader.ddspf.dwRBitMask = &HFF
                ddsHeader.ddspf.dwGBitMask = &HFF00
                ddsHeader.ddspf.dwBBitMask = &HFF0000
                ddsHeader.ddspf.dwABitMask = &HFF000000
                ddsHeader.dwCaps = &H1002
        End Select

        ddsHeader.dwDepth = 1

        If IsCubemap Then
            ddsHeader.dwCaps2 = &HFE00
        Else
            ddsHeader.dwCaps2 = 0
        End If

        If mipMapCount Then
            ddsHeader.dwMipMapCount = mipMapCount
        Else
            ddsHeader.dwMipMapCount = 1 + Math.Floor(Math.Log(Math.Max(ddsHeader.dwWidth, ddsHeader.dwHeight), 2))
        End If

        ddsHeader.ddspf.dwSize = &H20

        Return ddsHeader
    End Function

    Private Function MakeDDSDX10Header(dxgiFormat As UInteger, IsCubemap As Boolean) As DDS_HEADER_DXT10
        Dim ddsDx10Header As New DDS_HEADER_DXT10

        ddsDx10Header.dxgiFormat = dxgiFormat
        ddsDx10Header.resourceDimension = 3
        If IsCubemap Then
            ddsDx10Header.miscFlag = 4
            ddsDx10Header.arraySize = 6
        Else
            ddsDx10Header.miscFlag = 0
            ddsDx10Header.arraySize = 1
        End If
        ddsDx10Header.miscFlags2 = 0

        Return ddsDx10Header
    End Function

    Private Sub DeswizzleDDSBytes(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            DeswizzleDDSBytes(Width / 2, Height / 2, offset, offsetFactor * 2)
            DeswizzleDDSBytes(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            DeswizzleDDSBytes(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            DeswizzleDDSBytes(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            outBytes(offset * 3) = inBytes(writeOffset + 3)
            outBytes(offset * 3 + 1) = inBytes(writeOffset + 2)
            outBytes(offset * 3 + 2) = inBytes(writeOffset + 1)

            outBytes(offset * 3 + 3) = inBytes(writeOffset + 7)
            outBytes(offset * 3 + 4) = inBytes(writeOffset + 6)
            outBytes(offset * 3 + 5) = inBytes(writeOffset + 5)

            outBytes(offset * 3 + ddsWidth * 3) = inBytes(writeOffset + 11)
            outBytes(offset * 3 + ddsWidth * 3 + 1) = inBytes(writeOffset + 10)
            outBytes(offset * 3 + ddsWidth * 3 + 2) = inBytes(writeOffset + 9)

            outBytes(offset * 3 + ddsWidth * 3 + 3) = inBytes(writeOffset + 15)
            outBytes(offset * 3 + ddsWidth * 3 + 4) = inBytes(writeOffset + 14)
            outBytes(offset * 3 + ddsWidth * 3 + 5) = inBytes(writeOffset + 13)

            writeOffset += 16
        End If
    End Sub

    Private Sub SwizzleDDSBytes(Width As UShort, Height As UShort, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            SwizzleDDSBytes(Width / 2, Height / 2, offset, offsetFactor * 2)
            SwizzleDDSBytes(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            SwizzleDDSBytes(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            SwizzleDDSBytes(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            outBytes(writeOffset + 3) = inBytes(offset * 4)
            outBytes(writeOffset + 2) = inBytes(offset * 4 + 1)
            outBytes(writeOffset + 1) = inBytes(offset * 4 + 2)
            outBytes(writeOffset) = inBytes(offset * 4 + 3)

            outBytes(writeOffset + 7) = inBytes(offset * 4 + 4)
            outBytes(writeOffset + 6) = inBytes(offset * 4 + 5)
            outBytes(writeOffset + 5) = inBytes(offset * 4 + 6)
            outBytes(writeOffset + 4) = inBytes(offset * 4 + 7)

            outBytes(writeOffset + 11) = inBytes(offset * 4 + ddsWidth * 4)
            outBytes(writeOffset + 10) = inBytes(offset * 4 + ddsWidth * 4 + 1)
            outBytes(writeOffset + 9) = inBytes(offset * 4 + ddsWidth * 4 + 2)
            outBytes(writeOffset + 8) = inBytes(offset * 4 + ddsWidth * 4 + 3)

            outBytes(writeOffset + 15) = inBytes(offset * 4 + ddsWidth * 4 + 4)
            outBytes(writeOffset + 14) = inBytes(offset * 4 + ddsWidth * 4 + 5)
            outBytes(writeOffset + 13) = inBytes(offset * 4 + ddsWidth * 4 + 6)
            outBytes(writeOffset + 12) = inBytes(offset * 4 + ddsWidth * 4 + 7)

            writeOffset += 16
        End If
    End Sub

    'Private Sub DeswizzleDDSBytesPS4Uncompressed(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
    '    If (Width * Height > 4) Then
    '        DeswizzleDDSBytesPS4Uncompressed(Width / 2, Height / 2, offset, offsetFactor * 2)
    '        DeswizzleDDSBytesPS4Uncompressed(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
    '        DeswizzleDDSBytesPS4Uncompressed(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
    '        DeswizzleDDSBytesPS4Uncompressed(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
    '    Else
    '        If Not ddsBlockArray(offset * 8 / 16) And offset * 8 < outBytes.Length Then

    '            For i As UInteger = 0 To 15
    '                outBytes(offset * 8 + i) = inBytes(writeOffset + i)
    '            Next


    '            writeOffset += 16


    '            For i As UInteger = 0 To 15
    '                outBytes(offset * 8 + ddsWidth * 8 + i) = inBytes(writeOffset + i)
    '            Next
    '            ddsBlockArray(offset * 8 / 16) = True
    '            ddsBlockArray((offset * 8 + ddsWidth * 8) / 16) = True

    '            writeOffset += 16
    '        Else

    '            writeOffset += 32
    '        End If

    '    End If
    'End Sub

    Private Sub DeswizzleDDSBytesPS4RGBA(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            DeswizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset, offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            For i As UInteger = 0 To 15
                outBytes(offset * 8 + i) = inBytes(writeOffset + i)
            Next

            writeOffset += 16

            For i As UInteger = 0 To 15
                outBytes(offset * 8 + ddsWidth * 8 + i) = inBytes(writeOffset + i)
            Next

            writeOffset += 16
        End If
    End Sub

    Private Sub DeswizzleDDSBytesPS4RGBA8(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            DeswizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset, offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            DeswizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            For i As UInteger = 0 To 7
                outBytes(offset * 4 + i) = inBytes(writeOffset + i)
            Next

            writeOffset += 8

            For i As UInteger = 0 To 7
                outBytes(offset * 4 + ddsWidth * 4 + i) = inBytes(writeOffset + i)
            Next

            writeOffset += 8
        End If
    End Sub


    'Private Sub DeswizzleDDSBytesPS4(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger, format As UInteger)
    '    Dim BlockSize As UInteger
    '    Select Case (format)
    '        Case 0, 1, 25, 103, 108, 109
    '            BlockSize = 8
    '        Case 5, 100, 102, 106, 107, 110
    '            BlockSize = 16
    '    End Select

    '    If (Width * Height > 16) Then
    '        DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset, offsetFactor * 2, format)
    '        DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + (Width / 8) * BlockSize, offsetFactor * 2, format)
    '        If ddsWidth < 4 Then
    '            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + ((4 / 8) * (Height / 4) * BlockSize), offsetFactor * 2, format)
    '            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + (4 / 8) * (Height / 4) * BlockSize + (Width / 8) * BlockSize, offsetFactor * 2, format)
    '        Else
    '            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + ((ddsWidth / 8) * (Height / 4) * BlockSize), offsetFactor * 2, format)
    '            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + (ddsWidth / 8) * (Height / 4) * BlockSize + (Width / 8) * BlockSize, offsetFactor * 2, format)
    '        End If

    '    Else
    '        If (offset < outBytes.Length) Then
    '            If Not ddsBlockArray(offset / BlockSize) Then
    '                Dim ZeroCount = 0
    '                For i As UInteger = 0 To BlockSize - 1
    '                    If inBytes(writeOffset + i) = 0 Then
    '                        ZeroCount += 1
    '                    End If
    '                Next
    '                If ZeroCount <> BlockSize Then
    '                    For i As UInteger = 0 To BlockSize - 1
    '                        outBytes(offset + i) = inBytes(writeOffset)
    '                        writeOffset += 1
    '                    Next
    '                    ddsBlockArray(offset / BlockSize) = True
    '                Else
    '                    writeOffset += BlockSize
    '                End If

    '            Else
    '                writeOffset += BlockSize
    '            End If

    '        Else
    '            writeOffset += BlockSize
    '        End If


    '    End If
    'End Sub

    Private Sub DeswizzleDDSBytesPS4(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger, format As UInteger)
        Dim BlockSize As UInteger
        Select Case (format)
            Case 0, 1, 25, 103, 108, 109
                BlockSize = 8
            Case 5, 100, 102, 106, 107, 110
                BlockSize = 16
        End Select

        If (Width * Height > 16) Then
            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset, offsetFactor * 2, format)
            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + (Width / 8) * BlockSize, offsetFactor * 2, format)
            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + ((ddsWidth / 8) * (Height / 4) * BlockSize), offsetFactor * 2, format)
            DeswizzleDDSBytesPS4(Width / 2, Height / 2, offset + (ddsWidth / 8) * (Height / 4) * BlockSize + (Width / 8) * BlockSize, offsetFactor * 2, format)
        Else
            For i As UInteger = 0 To BlockSize - 1
                outBytes(offset + i) = inBytes(writeOffset)
                writeOffset += 1
            Next
        End If

    End Sub

    Private Sub DeswizzleDDSBytesPS4(Width As UInteger, Height As UInteger, Format As UInteger)
        Dim BlockSize As UInteger = 0
        Select Case (Format)
            Case 22
                BlockSize = 8
            Case 105
                BlockSize = 4
            Case 0, 1, 25, 103, 108, 109
                BlockSize = 8
            Case 5, 100, 102, 106, 107, 110
                BlockSize = 16
        End Select
        'If Format <> 22 And Width <= 4 And Height <= 4 Then
        '    For i As UInteger = 0 To BlockSize - 1
        '        outBytes(i) = inBytes(i)
        '    Next
        '    Return
        'End If
        'If Format = 22 And Width <= 2 And Height = 1 Then
        '    For i As UInteger = 0 To Width * 8 - 1
        '        outBytes(i) = inBytes(i)
        '    Next
        '    Return
        'End If
        Dim SwizzleBlockSize As UInteger = 0

        Dim BlocksH As UInteger
        Dim BlocksV As UInteger
        If Format = 22 Then
            BlocksH = (Width + 7) \ 8
            BlocksV = (Height + 7) \ 8
            SwizzleBlockSize = 8
        ElseIf Format = 105 Then
            BlocksH = (Width + 15) \ 16
            BlocksV = (Height + 15) \ 16
            SwizzleBlockSize = 16
        Else
            BlocksH = (Width + 31) \ 32
            BlocksV = (Height + 31) \ 32
            SwizzleBlockSize = 32
        End If

        Dim h As UInteger = 0
        Dim v As UInteger = 0
        Dim offset As UInteger = 0

        If Format = 22 Then
            DeswizzleDDSBytesPS4RGBA(Width, Height, 0, 2) 'not sure if this works for images larger than 8x8
            writeOffset = 0
            Return
        End If
        'ElseIf Format = 105 Then
        '    DeswizzleDDSBytesPS4RGBA8(Width, Height, 0, 2)
        '    writeOffset = 0
        '    Return
        'End If

        For i As UInteger = 0 To BlocksV - 1
            h = 0
            For j As UInteger = 0 To BlocksH - 1
                offset = h + v

                If Format = 105 Then
                    DeswizzleDDSBytesPS4RGBA8(16, 16, offset, 2)
                Else
                    DeswizzleDDSBytesPS4(32, 32, offset, 2, Format)
                End If


                h += (SwizzleBlockSize / 4) * BlockSize
                'SwizzleBlockSize = 32
            Next
            If Format = 105 Then
                v += SwizzleBlockSize * SwizzleBlockSize
            Else
                If BlockSize = 8 Then
                    v += SwizzleBlockSize * Width / 2
                Else
                    v += SwizzleBlockSize * Width
                End If
            End If
        Next

        writeOffset = 0
    End Sub

    Private Sub SwizzleDDSBytesPS4(Width As UInteger, Height As UInteger, Format As UInteger)
        Dim BlockSize As UInteger = 0
        Select Case (Format)
            Case 105
                BlockSize = 4
            Case 22
                BlockSize = 8
            Case 0, 1, 25, 103, 108, 109
                BlockSize = 8
            Case 5, 100, 102, 106, 107, 110
                BlockSize = 16
        End Select
        'If Format <> 22 And Width <= 4 And Height <= 4 Then
        '    For i As UInteger = 0 To BlockSize - 1
        '        outBytes(i) = inBytes(i)
        '    Next
        '    Return
        'End If
        'If Format = 22 And Width <= 2 And Height = 1 Then
        '    For i As UInteger = 0 To Width * 8 - 1
        '        outBytes(i) = inBytes(i)
        '    Next
        '    Return
        'End If
        Dim SwizzleBlockSize As UInteger = 0

        Dim BlocksH As UInteger
        Dim BlocksV As UInteger
        If Format = 22 Then
            BlocksH = (Width + 7) \ 8
            BlocksV = (Height + 7) \ 8
            SwizzleBlockSize = 8
        ElseIf Format = 105 Then
            BlocksH = (Width + 15) \ 16
            BlocksV = (Height + 15) \ 16
            SwizzleBlockSize = 16
        Else
            BlocksH = (Width + 31) \ 32
            BlocksV = (Height + 31) \ 32
            SwizzleBlockSize = 32
        End If

        Dim h As UInteger = 0
        Dim v As UInteger = 0
        Dim offset As UInteger = 0

        If Format = 22 Then
            SwizzleDDSBytesPS4RGBA(8, 8, 0, 2) 'not sure if this works for images larger than 8x8
            writeOffset = 0
            Return
        End If

        For i As UInteger = 0 To BlocksV - 1
            h = 0
            For j As UInteger = 0 To BlocksH - 1
                offset = h + v

                If Format = 105 Then
                    SwizzleDDSBytesPS4RGBA8(16, 16, offset, 2)
                Else
                    SwizzleDDSBytesPS4(32, 32, offset, 2, Format)
                End If


                h += (SwizzleBlockSize / 4) * BlockSize
                'SwizzleBlockSize = 32
            Next

            If BlockSize = 8 Then
                v += SwizzleBlockSize * Width / 2
            Else
                v += SwizzleBlockSize * Width
            End If
        Next

        writeOffset = 0
    End Sub

    Private Sub SwizzleDDSBytesPS4(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger, format As UInteger)
        Dim BlockSize As UInteger
        Select Case (format)
            Case 0, 1, 25, 103, 108, 109
                BlockSize = 8
            Case 5, 100, 102, 106, 107, 110
                BlockSize = 16
        End Select

        If (Width * Height > 16) Then
            SwizzleDDSBytesPS4(Width / 2, Height / 2, offset, offsetFactor * 2, format)
            SwizzleDDSBytesPS4(Width / 2, Height / 2, offset + (Width / 8) * BlockSize, offsetFactor * 2, format)
            SwizzleDDSBytesPS4(Width / 2, Height / 2, offset + ((ddsWidth / 8) * (Height / 4) * BlockSize), offsetFactor * 2, format)
            SwizzleDDSBytesPS4(Width / 2, Height / 2, offset + (ddsWidth / 8) * (Height / 4) * BlockSize + (Width / 8) * BlockSize, offsetFactor * 2, format)
        Else
            For i As UInteger = 0 To BlockSize - 1
                outBytes(writeOffset) = inBytes(offset + i)
                writeOffset += 1
            Next
        End If
    End Sub

    Private Sub SwizzleDDSBytesPS4RGBA(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            SwizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset, offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            For i As UInteger = 0 To 15
                outBytes(writeOffset + i) = inBytes(offset * 8 + i)
            Next

            writeOffset += 16

            For i As UInteger = 0 To 15
                outBytes(writeOffset + i) = inBytes(offset * 8 + ddsWidth * 8 + i)
            Next

            writeOffset += 16
        End If
    End Sub

    Private Sub SwizzleDDSBytesPS4RGBA8(Width As UInteger, Height As UInteger, offset As UInteger, offsetFactor As UInteger)
        If (Width * Height > 4) Then
            SwizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset, offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + (Width / 2), offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor), offsetFactor * 2)
            SwizzleDDSBytesPS4RGBA8(Width / 2, Height / 2, offset + ((Width / 2) * (Height / 2) * offsetFactor) + (Width / 2), offsetFactor * 2)
        Else
            For i As UInteger = 0 To 7
                outBytes(writeOffset + i) = inBytes(offset * 4 + i)
            Next

            writeOffset += 8

            For i As UInteger = 0 To 7
                outBytes(writeOffset + i) = inBytes(offset * 4 + ddsWidth * 4 + i)
            Next

            writeOffset += 8
        End If
    End Sub

    Private Sub ExtendInBytes(IsCubemap As Boolean)
        Dim tempBytes(inBytes.Length * (4 / 3) - 1) As Byte
        For i As UInteger = 0 To inBytes.Length / 3 - 1
            tempBytes(i * 4) = inBytes(i * 3)
            tempBytes(i * 4 + 1) = inBytes(i * 3 + 1)
            tempBytes(i * 4 + 2) = inBytes(i * 3 + 2)
            If IsCubemap Then
                tempBytes(i * 4 + 3) = &HFF
            Else
                If i = inBytes.Length / 3 - 1 Then
                    tempBytes(i * 4 + 3) = &H11
                Else
                    tempBytes(i * 4 + 3) = inBytes((i + 1) * 3)
                End If
            End If
        Next
        ReDim inBytes(tempBytes.Length - 1)
        Array.Copy(tempBytes, inBytes, tempBytes.Length)
    End Sub

    Private Sub AddMemeBytes(ByRef ddsBytes As Byte())
        ReDim Preserve ddsBytes(ddsBytes.Length + &H48 * 5 - 1)
        Dim memeBytes(&H47) As Byte
        For i As UInteger = 0 To &H47
            memeBytes(i) = &HEE
        Next
        For i As UInteger = 0 To 4
            Array.Copy(ddsBytes, &HAB8 + (4 - i) * &HAB8, ddsBytes, &HB00 + (4 - i) * &HB00, &HAB8)
            Array.Copy(memeBytes, 0, ddsBytes, &HAB8 + (4 - i) * &HB00, &H48)
        Next
    End Sub

    Private Sub RemoveMemeBytes(ByRef ddsBytes As Byte())
        For i As UInteger = 0 To 4
            Array.Copy(ddsBytes, &HB00 + i * &HB00, ddsBytes, &HAB8 + i * &HAB8, &HAB8)
        Next
        ReDim Preserve ddsBytes(ddsBytes.Length - &H48 * 5 - 1)
    End Sub

    Private Sub txt_Drop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        sender.Lines = file
    End Sub
    Private Sub txt_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub

    Private Sub Des_BNDBuild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        updateUITimer.Interval = 200
        updateUITimer.Start()
    End Sub

    Private Sub updateUI() Handles updateUITimer.Tick
        SyncLock workLock
            If work Then
                btnExtract.Enabled = False
                btnRebuild.Enabled = False
            Else
                btnExtract.Enabled = True
                btnRebuild.Enabled = True
            End If
        End SyncLock

        SyncLock outputLock
            While outputList.Count > 0
                txtInfo.AppendText(outputList(0))
                outputList.RemoveAt(0)
            End While
        End SyncLock


        If txtInfo.Lines.Count > 10000 Then
            Dim newList As List(Of String) = txtInfo.Lines.ToList
            While newList.Count > 1000
                newList.RemoveAt(0)
            End While
            txtInfo.Lines = newList.ToArray
        End If
    End Sub
End Class

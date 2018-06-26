Imports System.IO
Module Module1

    Const VERSION = "1.2"

    Const Target_Location = ""

    Class Hangul
        Public Enum charState
            기타 = 0
            자음 = 1
            모음 = 2
            음절 = 3

            초성 = 10
            종성 = 11
            중성 = 20

        End Enum

        Sub New(ByVal val As Integer, ByVal State As charState)
            Me.value = val
            Me.State = State

        End Sub

        Dim value As Integer = Nothing
        Dim State As charState = Nothing

        Public Function getValue() As Integer
            Return value
        End Function
        Public Function getChar() As Char
            Return ChrW(value)
        End Function
        Public Function getState() As charState
            Return State
        End Function

        Public Function combine(ByVal v1 As Integer, ByVal v2 As Integer, ByVal v3 As Integer) As Char
            Dim temp(3) As Integer
            temp(0) = (v1 - &H1100) * 21 * 28
            temp(1) = (v2 - &H1161) * 28
            temp(2) = (v3 - &H11A8) + 1
            'MsgBox(temp(0) + temp(1) + temp(2) + &HAC00)
            Return ChrW(temp(0) + temp(1) + temp(2) + &HAC00)
        End Function

    End Class

    Sub Main()

        Dim FileList As New List(Of String)
        Dim fileName As String = Nothing
        'Dim index As Integer = 0
        Dim BASE_LOCATION As String = Nothing
        Dim errorCount As Integer = 0
        Dim noChangeCount As Integer = 0
        Dim ChangeCount As Integer = 0
        Dim log_file As String = My.Application.Info.DirectoryPath & $"\Result_{Today.Year & Today.Month & Today.Day & Today.Hour & Today.Minute & Today.Second}.log"
        Console.WriteLine("---------------------------------------------------------------------------------------------------------------")
        Console.WriteLine(" 본 프로그램은 Adobe사의 애프터이펙트 프리셋 파일 이름이 깨지는 현상을 고치는 프로그램입니다.")
        Console.WriteLine(" Adobe의 공식 프로그램이 아니며, 아마추어 개발자가 작성한 프로그램입니다.")
        Console.WriteLine(" 본 프로그램은 타인에 의해 인증되지 않았기 때문에 모든 파일 이름이 고쳐질 것이란 보장은 없습니다.")
        Console.WriteLine(" 사용 전 프리셋 폴더 백업을 권장합니다.")
        Console.WriteLine(" 존재하지 않는 폴더 경로를 입력하면 다시 입력하라고 합니다.")
        Console.WriteLine(" 기본 경로는 C:\Program Files\Adobe After Effects CC 2018\Support Files\Presets 입니다.")
        Console.WriteLine()
        Console.WriteLine(" 지속적으로 업데이트 되므로 http://pang2h.tistory.com/49 에서 정식으로 다운받으시기 바랍니다.")
        Console.WriteLine()
        Console.WriteLine(" 사용법은 간단합니다. 프리셋 폴더의 경로(프리셋 폴더 이름 포함)를 입력후 Enter를 누르면 됩니다.")
        Console.WriteLine(" 작업 완료후 프로그램이 종료되면 결과창이 뜹니다. 오류를 찾아보려면 [ 오류 발생 ]을 검색하세요. (대괄호 제외)")
        Console.WriteLine("                                                                             Version : {0}", VERSION)
        Console.WriteLine("---------------------------------------------------------------------------------------------------------------")

        Do While (Not My.Computer.FileSystem.DirectoryExists(BASE_LOCATION))
            Console.Write(" 프리셋 폴더 경로 : ")
            BASE_LOCATION = Console.ReadLine()

        Loop

        For Each pos As String In My.Computer.FileSystem.GetFiles(BASE_LOCATION, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
            FileList.Add(pos)
            Console.WriteLine($" 대상 : {pos}")
            My.Computer.FileSystem.WriteAllText(log_file, $"대상 : {pos}{Chr(10) & Chr(13) & Chr(10) & Chr(13)}", True)
            Dim file As FileInfo = My.Computer.FileSystem.GetFileInfo(pos)
            fileName = getFileName(file.Name, file.Extension)
            Dim D() As Hangul = isHangulExists(fileName)
            fileName = Nothing
            Dim S(3) As Integer
            Dim Last As Hangul.charState = 0
            For i = 0 To D.Count - 1
                Select Case (D(i).getState)
                    Case Hangul.charState.초성
                        If Last = Hangul.charState.중성 Then
                            fileName &= D(i).combine(S(0), S(1), 4519)
                        End If
                        S(0) = D(i).getValue
                        Last = Hangul.charState.초성
                    Case Hangul.charState.중성
                        S(1) = D(i).getValue
                        If i + 1 = D.Count Then
                            fileName &= D(i).combine(S(0), S(1), 4519)
                        End If
                        Last = Hangul.charState.중성
                    Case Hangul.charState.종성
                        S(2) = D(i).getValue
                        fileName &= D(i).combine(S(0), S(1), S(2))
                        Last = Hangul.charState.종성
                    Case Else
                        If Last = Hangul.charState.중성 Then
                            fileName &= D(i).combine(S(0), S(1), 4519)

                        End If
                        fileName &= D(i).getChar
                        Last = Hangul.charState.기타
                End Select
            Next
            Last = Nothing
            If My.Computer.FileSystem.FileExists($"{file.DirectoryName}\{fileName}{file.Extension}") Then
                Console.WriteLine($" {Chr(9)}>> 변경 안함")
                Console.WriteLine()
                My.Computer.FileSystem.WriteAllText(log_file, $"{Chr(9)}>> 변경 안함{Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) }", True)
                noChangeCount += 1
            Else
                Try
                    My.Computer.FileSystem.RenameFile(pos, $"{fileName}{file.Extension}")
                    Console.WriteLine($" {Chr(9)}>> { file.DirectoryName}\{fileName & file.Extension}")
                    Console.WriteLine()
                    My.Computer.FileSystem.WriteAllText(log_file, $"{Chr(9)}>> { file.Directory}\{fileName & file.Extension}{Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) }", True)
                    ChangeCount += 1
                Catch ex As Exception
                    Console.WriteLine($" {Chr(9)}>> { file.DirectoryName}\{fileName & file.Extension}")
                    Console.WriteLine()
                    My.Computer.FileSystem.WriteAllText(log_file, $"{Chr(9)}>> 오류 발생{Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) }", True)
                    errorCount += 1
                End Try

            End If

        Next
        My.Computer.FileSystem.WriteAllText(log_file, $"전체 항목 수 : {noChangeCount + ChangeCount + errorCount}{Chr(9)}변경 수 : {ChangeCount}{Chr(9)}미변경 수 : {noChangeCount}{Chr(9)}오류 수 : {errorCount}", True)
        Console.WriteLine()
        Console.WriteLine(" 작업이 완료되었습니다, 프로그램을 끝내려면 아무 키나 누르세요.")
        Console.ReadKey()

        Process.Start(log_file)

    End Sub

    Function getFileName(ByVal value As String, ByVal Extension As String) As String
        Return Mid(value, 1, InStr(value, Extension) - 1)
    End Function

    Function isHangulExists(ByVal value As String) As Hangul()

        Dim arr As New List(Of Integer)
        Dim str(value.Count - 1) As Hangul
        'Console.WriteLine()
        'Console.Write(value & $": {value.Count}")
        For i = 0 To value.Length - 1
            arr.Add(AscW(value(i)))
        Next
        'Console.Write($", {arr.Count}")
        'Console.WriteLine()
        For i = 0 To arr.Count - 1
            'Console.Write(arr(i) & " ")
            Select Case arr(i)
                Case 4352 To 4371
                    str(i) = New Hangul(arr(i), Hangul.charState.초성)

                Case 4449 To 4470
                    str(i) = New Hangul(arr(i), Hangul.charState.중성)

                Case 4520 To 4546
                    str(i) = New Hangul(arr(i), Hangul.charState.종성)

                Case &HAC00 To &HD7AF
                    str(i) = New Hangul(arr(i), Hangul.charState.음절)

                Case Else
                    str(i) = New Hangul(arr(i), Hangul.charState.기타)

            End Select
        Next

        Return str
    End Function

End Module

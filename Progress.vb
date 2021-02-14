Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Web

Namespace FP930
    Public Class Progress

        Public Sub Add(ByVal [step] As ProgressStep)
            For Each s In Steps
                If s.Status <> ProgressStatus.[Error] Then
                    s.Status = ProgressStatus.Completed
                End If
            Next

            Steps.Add([step])
        End Sub

        Private _Steps As List(Of ProgressStep)

        Public Property Steps As List(Of ProgressStep)
            Get
                If _Steps Is Nothing Then
                    _Steps = New List(Of ProgressStep)()
                End If
                Return _Steps
            End Get
            Set(ByVal value As List(Of ProgressStep))
                _Steps = value
            End Set
        End Property

        'Public Overrides Function ToString() As String
        '    Dim sb As StringBuilder = New StringBuilder()
        '    sb.Append("<table>")

        '    For Each [step] In Steps
        '        sb.Append(String.Format("<tr><td>{0}&nbsp;&nbsp;</td><td>{1}&nbsp;&nbsp;-</td><td>{2}</td></tr>", If([step].Status = ProgressStatus.Completed, "<img src='icons/check.ico' alt='' height='20px' />", If([step].Status = ProgressStatus.InProgress, "<img src='icons/loading.gif' alt='' height='20px' />", If([step].Status = ProgressStatus.[Error], "<img src='icons/error.png' alt='' height='20px' />", "UNKNOWN"))), [step].StartTime.ToString("hh:mm tt"), [step].Message))
        '    Next

        '    sb.Append("</table>")
        '    Return sb.ToString()
        'End Function

        Public Overrides Function ToString() As String
            Dim sb As StringBuilder = New StringBuilder()

            Dim [step] As ProgressStep = Steps.LastOrDefault()
            sb.Append([step].StartTime.ToString("hh:mm tt") & " : " & [step].Message)

            Return sb.ToString()
        End Function

        Public Sub Dispose()
            GC.Collect()
        End Sub
    End Class

    Public Enum ProgressStatus
        [Error] = 100
        InProgress
        Completed
    End Enum

    Public Class ProgressStep
        Public Sub New(ByVal msg As String, ByVal status As ProgressStatus, ByVal Optional flag As String = "")
            If flag <> "" Then
                flag = GetHiddenField(flag)
            ElseIf status = ProgressStatus.[Error] Then
                flag = GetHiddenField("000ABORT_CHECK000")
            End If

            StartTime = DateTime.Now
            Message = msg & flag
            status = status
        End Sub

        Public Property Message As String
        Public Property StartTime As DateTime
        Public Property Status As ProgressStatus

        Private Function GetHiddenField(ByVal flag As String) As String
            'Return "<div hidden='hidden'>" & flag & "</div>"
            Return ""
        End Function
    End Class
End Namespace

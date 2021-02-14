'********
'**     FG035 : Eternal Report Handler class
'**             git testing
'********

Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading.Tasks
Imports FP901.Eternal.Asp.Software

Namespace Eternal.Asp.Appl
    Public Class erep

        '**************************************************************************************************
        '**                                           Globals                                            **
        '*************************************************************************************************

        Private l035_cosrl As Short                             '** Company Serial                                              
        Private l035_finyr As Short                             '** Financial Year     
        Private l035_reportid As Integer
        Private l035_userid As Integer
        Private l035_reqdate As DateTime
        Private l035_pagecd As String
        Private l035_repcode As String
        Private l035_uco As String
        Private l035_param1 As String
        Private l035_param2 As String
        Private l035_param3 As String
        Private l035_param4 As String
        Private l035_reqstatus As String
        Private l035_outputtype As String
        Private l035_genstatus As String
        Private l035_genstart As DateTime
        Private l035_genend As DateTime
        Private l035_repfile As String
        Private l035_repdir As String
        Private l035_filesize As Long
        Private l035_timetaken As Long
        Private l035_trycount As Integer
        Private l035_errormsg As String
        Private l035_adc As String

        Private l035_repno As String
        Private l035_paramlist As New List(Of param)
        Private l035_runwith As String
        Private l035_lds As DataSet

        Private l035_conn As SqlConnection
        Private l035_cmd As SqlCommand


        '**************************************************************************************************
        '**                                      Public Functions                                        **
        '**************************************************************************************************

        '*******
        '**    This would move default values to global fields.
        '*******
        Public Sub New()
            Call FG035_initme()
        End Sub

        '********
        '**     FT035_initme() : This will initialize the globals.
        '********
        Public Sub FG035_initme()

            l035_cosrl = CShort(asplib.pickint("gi_cosrl"))
            l035_finyr = CShort(asplib.pickint("gi_finyr"))
            l035_reportid = 0
            l035_userid = 0
            l035_reqdate = Date.Now
            l035_pagecd = ""
            l035_repcode = ""
            l035_uco = ""
            l035_param1 = ""
            l035_param2 = ""
            l035_param3 = ""
            l035_param4 = ""
            l035_reqstatus = ""
            l035_outputtype = ""
            l035_genstatus = ""
            l035_genstart = Date.Now
            l035_genend = Date.Now
            l035_repdir = ""
            l035_repfile = ""
            l035_filesize = 0
            l035_timetaken = 0
            l035_trycount = 0
            l035_errormsg = ""
            l035_adc = "X"

            l035_repno = ""
            l035_paramlist.Clear()
            l035_runwith = ""
        End Sub


        '*******
        '**  update : This function is used to add/update record in Z710 table.
        '*******
        Public Function update() As Boolean
            Dim lb_ret As Boolean
            Dim lc_sql As String
            Dim lc_field As String
            Dim lc_value As String

            lb_ret = False

            Try

                Select Case l035_adc.ToUpper
                    Case "A"
                        l035_reportid = DBLIB.getnewid("z710", "reportid")

                        lc_field = "Insert into Z710 ( "
                        lc_value = " Values ("
                        lc_field = lc_field & " Z710_cosrl"
                        lc_value = lc_value & " " & l035_cosrl & " "
                        lc_field = lc_field & ", Z710_finyr"
                        lc_value = lc_value & ", " & l035_finyr & " "
                        lc_field = lc_field & ",Z710_reportid"
                        lc_value = lc_value & ", " & l035_reportid & " "
                        lc_field = lc_field & ",Z710_userid"
                        lc_value = lc_value & ", " & l035_userid & " "
                        lc_field = lc_field & ",Z710_reqdate"
                        lc_value = lc_value & ", '" & Format(l035_reqdate, "MM/dd/yyyy hh:mm:ss") & "' "
                        lc_field = lc_field & ", Z710_pagecd"
                        lc_value = lc_value & ", '" & l035_pagecd & "' "
                        lc_field = lc_field & ", Z710_repcode"
                        lc_value = lc_value & ", '" & l035_repcode & "' "
                        lc_field = lc_field & ",Z710_uco "
                        lc_value = lc_value & ", '" & l035_uco & "' "
                        lc_field = lc_field & ",Z710_param1 "
                        lc_value = lc_value & ", '" & l035_param1 & "' "
                        lc_field = lc_field & ",Z710_param2 "
                        lc_value = lc_value & ", '" & l035_param2 & "' "
                        lc_field = lc_field & ",Z710_param3 "
                        lc_value = lc_value & ", '" & l035_param3 & "' "
                        lc_field = lc_field & ",Z710_param4 "
                        lc_value = lc_value & ", '" & l035_param4 & "' "
                        lc_field = lc_field & ",Z710_reqstatus "
                        lc_value = lc_value & ", '" & l035_reqstatus & "' "
                        lc_field = lc_field & ",Z710_outputtype "
                        lc_value = lc_value & ", '" & l035_outputtype & "' "
                        lc_field = lc_field & ",Z710_genstatus "
                        lc_value = lc_value & ", '" & l035_genstatus & "' "
                        lc_field = lc_field & ",Z710_genstart "
                        lc_value = lc_value & ", '" & Format(l035_genstart, "MM/dd/yyyy hh:mm:ss") & "' "
                        lc_field = lc_field & ",Z710_genend "
                        lc_value = lc_value & ", '" & Format(l035_genend, "MM/dd/yyyy hh:mm:ss") & "' "
                        lc_field = lc_field & ",Z710_repfile "
                        lc_value = lc_value & ", '" & l035_repfile & "' "
                        lc_field = lc_field & ",Z710_repdir "
                        lc_value = lc_value & ", '" & l035_repdir & "' "
                        lc_field = lc_field & ",Z710_filesize"
                        lc_value = lc_value & ", " & l035_filesize & " "
                        lc_field = lc_field & ",Z710_timetaken"
                        lc_value = lc_value & ", " & l035_timetaken & " "
                        lc_field = lc_field & ",Z710_trycount )"
                        lc_value = lc_value & ", " & l035_trycount & ") "

                        lc_sql = lc_field & lc_value
                        lb_ret = DBLIB.update(lc_sql)

                    Case "C"
                        lc_sql = "Update z710 set "
                        lc_sql = lc_sql & "  Z710_reqstatus = '" & l035_reqstatus & "'"
                        lc_sql = lc_sql & ", Z710_outputtype = '" & l035_outputtype & "'"
                        lc_sql = lc_sql & ", Z710_genstatus = '" & l035_genstatus & "'"
                        lc_sql = lc_sql & ", Z710_genend = '" & Format(l035_genend, "MM/dd/yyyy hh:mm:ss") & "' "
                        lc_sql = lc_sql & ", Z710_repfile = '" & l035_repfile & "'"
                        lc_sql = lc_sql & ", Z710_repdir = '" & l035_repdir & "'"
                        lc_sql = lc_sql & ", Z710_filesize = " & l035_filesize & ""
                        lc_sql = lc_sql & ", Z710_timetaken = " & l035_timetaken & ""
                        lc_sql = lc_sql & ", Z710_trycount = " & l035_trycount & ""
                        lc_sql = lc_sql & " where Z710_reportid = " & l035_reportid

                        lb_ret = DBLIB.update(lc_sql)
                End Select

            Catch ex As Exception
                update = False
                l035_errormsg = "FG035-update : " & ex.Message
            End Try

FG035_exit:
            update = lb_ret
        End Function


        '*********
        '** param : This will make the collection of given parameters using the param class.
        '*********
        Public Sub param(ByVal pi_paramno As Short, ByVal pc_desc As String, ByVal po_value As Object)
            Dim lo_param As New param

            lo_param.paramno = pi_paramno
            lo_param.desc = pc_desc.Trim
            lo_param.value = po_value
            l035_paramlist.Add(lo_param)

        End Sub


        '*********
        '** param : This will return param object for the given parameter nnumber.
        '*********
        Public Function param(ByVal pi_paramno As Short) As param
            Dim lo_param As param

            lo_param = l035_paramlist.FirstOrDefault(Function(x) x.paramno = pi_paramno)
            param = lo_param
        End Function


        '**********
        '** startme : 
        '**********
        Public Function startme() As Boolean
            Dim lb_ret As Boolean

            lb_ret = FG035_validate()
            If (lb_ret = False) Then
                GoTo FG035_exit
            End If

            asplib.putobj("goa_erep", Me)

FG035_exit:
            startme = lb_ret
        End Function


        '********
        '** FG035_validate : This will validate the given input parameters. If everything is ok
        '** it will return true else false.
        '*********
        Private Function FG035_validate() As Boolean
            l035_errormsg = ""
            Return True
        End Function


        '***********
        '** registerme : This will add the report record in z710 table after all the parameters have
        '**              been validated.
        '***********
        Public Function registerme() As Boolean
            Dim lb_ret As Boolean

            l035_userid = asplib.pickint("gi0_userid")
            l035_reqdate = Date.Now
            l035_pagecd = asplib.pickalpha("gca_callingpage")
            l035_repcode = l035_repno
            l035_uco = l035_repno & "0"
            l035_param1 = If(param(1) Is Nothing, "", CType(param(1).value, String))
            l035_param2 = If(param(2) Is Nothing, "", CType(param(2).value, String))
            l035_param3 = If(param(3) Is Nothing, "", CType(param(3).value, String))
            l035_param4 = If(param(4) Is Nothing, "", CType(param(4).value, String))
            l035_reqstatus = "P"
            l035_outputtype = "P"
            l035_genstatus = "P"
            l035_genstart = Date.Now
            l035_genend = Date.Now
            l035_repfile = ""
            l035_repdir = ""
            l035_filesize = 0
            l035_timetaken = 0
            l035_trycount = 0
            l035_adc = "A"

            lb_ret = update()

FG035_exit:
            registerme = lb_ret
        End Function


        '********
        '** getdata : This will retrives the data for the given input parameters.
        '********
        Public Function getdata() As Boolean
            Dim lb_ret As Boolean

            lb_ret = True
            getdata = lb_ret
        End Function


        '********
        '** runme : This will actually handle the code for the report generation.
        '********
        Public Function runme() As Boolean ' ByVal po_httpstate As Object
            Dim lo_replog As New replog
            Dim lb_ret As Boolean
            Dim lc_repfile As String

            'Dim lo_httpstate As httpstate = TryCast(po_httpstate, httpstate)

            'If lo_httpstate IsNot Nothing AndAlso lo_httpstate.httpcontext IsNot Nothing Then
            '    HttpContext.Current = lo_httpstate.httpcontext
            'End If

            Dim lc_sqledition As String
            Dim lc_dbpass As String
            Dim lc_applserver As String
            Dim lc_str As String

            lc_sqledition = Trim(CStr(asplib.pickalpha("gc1_sqledition")))
            lc_dbpass = Trim(CStr(asplib.pickalpha("gc1_dbpass")))
            lc_applserver = Trim(CStr(asplib.pickalpha("gc_applserver")))

            If (LCase(lc_sqledition) = "reg") Then
                lc_str = lc_applserver.Trim
                If (LCase(lc_dbpass) = "dv") Then
                    lc_applserver = lc_str & ";uid=kt_dev;pwd=kt_dev"
                Else
                    lc_applserver = lc_str & ";uid=kt_prod;pwd=kt_prd1"
                End If
            End If

            l035_conn = New SqlConnection(lc_applserver)


            '*******
            '** Add report log in z711 table using replog(FG036-FP930) class runme fn
            '*******
            For li As Integer = 1 To 10
                lb_ret = FG035_addlog()
                If (lb_ret = False) Then
                    GoTo FG035_exit
                End If

                System.Threading.Thread.Sleep(1000)
            Next

            lc_repfile = "R4444-print.pdf"

            Dim lo_fileinfo As New FileInfo(HttpContext.Current.Server.MapPath("~/Reports/" & lc_repfile))

            l035_reqstatus = "T"
            l035_outputtype = "P"
            l035_genstatus = "G"
            l035_genend = Date.Now
            l035_repfile = lc_repfile
            l035_repdir = HttpContext.Current.Server.MapPath("~/Reports")
            l035_filesize = lo_fileinfo.Length
            l035_timetaken = 5000
            l035_trycount = 1
            l035_adc = "C"

            lb_ret = FG035_finalupdate()
            If (lb_ret = False) Then
                GoTo FG035_exit
            End If

FG035_exit:
            runme = lb_ret
        End Function



        '*********
        '** FG035_addlog : This will log record in z711 table for the given report id.
        '*********
        Private Function FG035_addlog() As Boolean
            Dim lc_sql As String
            Dim lc_field As String
            Dim lc_value As String
            Dim lb_ret As Boolean
            Dim li_srlno As Integer
            Dim li_count As Integer


            Try

                li_srlno = FG035_getnextsrl()

                lc_field = "Insert into z711 ( "
                lc_value = " Values ("
                lc_field = lc_field & " z711_reportid"
                lc_value = lc_value & " " & l035_reportid & " "
                lc_field = lc_field & ", z711_srlno"
                lc_value = lc_value & ", " & li_srlno & " "
                lc_field = lc_field & ",z711_desc"
                lc_value = lc_value & ", '" & "Description " & li_srlno & "' "
                lc_field = lc_field & ",z711_updatetime )"
                lc_value = lc_value & ", '" & Format(Date.Now, "MM/dd/yyyy hh:mm:ss") & "' )"

                lc_sql = lc_field & lc_value

                l035_cmd = New SqlCommand(lc_sql, l035_conn)
                l035_conn.Open()
                li_count = l035_cmd.ExecuteNonQuery()
                If (li_count > 0) Then
                    lb_ret = True
                Else
                    lb_ret = False
                End If

            Catch ex As Exception
                lb_ret = False
                l035_errormsg = "FG035-addlog : Exception-" & ex.Message
            Finally
                l035_conn.Close()
                l035_cmd.Parameters.Clear()
            End Try

            FG035_addlog = lb_ret
        End Function

        '**************
        '** FG035_getnextsrl : This will find the next serial number for the given report.
        '**************
        Private Function FG035_getnextsrl() As Integer
            Dim li_srlno As Integer
            Dim lc_sql As String
            Dim lds As New DataSet
            Dim lo_da As New SqlDataAdapter

            lc_sql = "select z711_srlno from z711 where " &
                     " ( z711_reportid =" & l035_reportid & ") " &
                     " order by z711_srlno desc"

            l035_cmd = New SqlCommand(lc_sql, l035_conn)
            lo_da.SelectCommand = l035_cmd
            lo_da.Fill(lds)
            If (DBLIB.hasdata(lds) = False) Then
                li_srlno = 1
            Else
                li_srlno = CInt(lds.Tables(0).Rows(0)("z711_srlno")) + 1
            End If

            FG035_getnextsrl = li_srlno
        End Function



        '**********
        '** FG035_finalupdate : This will update the z711 table with status that the current report is generated.
        '**********
        Private Function FG035_finalupdate() As Boolean
            Dim lc_sql As String
            Dim lb_ret As Boolean
            Dim li_count As Integer

            lc_sql = "Update z710 set "
            lc_sql = lc_sql & "  Z710_reqstatus = '" & l035_reqstatus & "'"
            lc_sql = lc_sql & ", Z710_outputtype = '" & l035_outputtype & "'"
            lc_sql = lc_sql & ", Z710_genstatus = '" & l035_genstatus & "'"
            lc_sql = lc_sql & ", Z710_genend = '" & Format(l035_genend, "MM/dd/yyyy hh:mm:ss") & "' "
            lc_sql = lc_sql & ", Z710_repfile = '" & l035_repfile & "'"
            lc_sql = lc_sql & ", Z710_repdir = '" & l035_repdir & "'"
            lc_sql = lc_sql & ", Z710_filesize = " & l035_filesize & ""
            lc_sql = lc_sql & ", Z710_timetaken = " & l035_timetaken & ""
            lc_sql = lc_sql & ", Z710_trycount = " & l035_trycount & ""
            lc_sql = lc_sql & " where Z710_reportid = " & l035_reportid

            Try
                l035_cmd = New SqlCommand(lc_sql, l035_conn)
                l035_conn.Open()
                li_count = l035_cmd.ExecuteNonQuery()
                If (li_count > 0) Then
                    lb_ret = True
                Else
                    lb_ret = False
                End If
            Catch ex As Exception
                lb_ret = False
                l035_errormsg = "FG035_finalupdate : Exception-" & ex.Message
            Finally
                l035_conn.Close()
                l035_cmd.Parameters.Clear()
            End Try
            FG035_finalupdate = lb_ret
        End Function





        '***************************************************************************************************
        '**                                       Get-Set Properties                                         **
        '***************************************************************************************************
        Public Property cosrl() As Short       '                                                                   
            Get
                cosrl = l035_cosrl
            End Get
            Set(ByVal pi_cosrl As Short)
                l035_cosrl = pi_cosrl
            End Set
        End Property

        Public Property finyr() As Short       '                                                                   
            Get
                finyr = l035_finyr
            End Get
            Set(ByVal pi_finyr As Short)
                l035_finyr = pi_finyr
            End Set
        End Property

        Public Property reportid() As Integer
            Get
                Return l035_reportid
            End Get
            Set(ByVal pi_reportid As Integer)
                l035_reportid = pi_reportid
            End Set
        End Property

        Public Property userid() As Integer       '                                                                   
            Get
                userid = l035_userid
            End Get
            Set(ByVal pi_userid As Integer)
                l035_userid = pi_userid
            End Set
        End Property
        Public Property reqdate() As DateTime        '                                                                   
            Get
                reqdate = l035_reqdate
            End Get
            Set(ByVal pd_reqdate As DateTime)
                l035_reqdate = pd_reqdate
            End Set
        End Property
        Public Property pagecd() As String        '                                                                   
            Get
                pagecd = l035_pagecd
            End Get
            Set(ByVal pc_pagecd As String)
                l035_pagecd = pc_pagecd
            End Set
        End Property

        Public Property repcode() As String
            Get
                Return l035_repcode
            End Get
            Set(ByVal pc_repcode As String)
                l035_repcode = pc_repcode
            End Set
        End Property

        Public Property uco() As String        '                                                                   
            Get
                uco = l035_uco
            End Get
            Set(ByVal pc_uco As String)
                l035_uco = pc_uco
            End Set
        End Property
        Public Property param1() As String
            Get
                param1 = l035_param1
            End Get
            Set(ByVal pc_param1 As String)
                l035_param1 = pc_param1
            End Set
        End Property
        Public Property param2() As String
            Get
                param2 = l035_param2
            End Get
            Set(ByVal pc_param2 As String)
                l035_param2 = pc_param2
            End Set
        End Property
        Public Property param3() As String
            Get
                param3 = l035_param3
            End Get
            Set(ByVal pc_param3 As String)
                l035_param3 = pc_param3
            End Set
        End Property
        Public Property param4() As String
            Get
                param4 = l035_param4
            End Get
            Set(ByVal pc_param4 As String)
                l035_param4 = pc_param4
            End Set
        End Property
        Public Property reqstatus() As String
            Get
                reqstatus = l035_reqstatus
            End Get
            Set(ByVal pc_reqstatus As String)
                l035_reqstatus = pc_reqstatus
            End Set
        End Property
        Public Property outputtype() As String
            Get
                outputtype = l035_outputtype
            End Get
            Set(ByVal pc_outputtype As String)
                l035_outputtype = pc_outputtype
            End Set
        End Property
        Public Property genstatus() As String
            Get
                genstatus = l035_genstatus
            End Get
            Set(ByVal pc_genstatus As String)
                l035_genstatus = pc_genstatus
            End Set
        End Property
        Public Property genstart() As DateTime
            Get
                genstart = l035_genstart
            End Get
            Set(ByVal pd_genstart As DateTime)
                l035_genstart = pd_genstart
            End Set
        End Property
        Public Property genend() As DateTime
            Get
                genend = l035_genend
            End Get
            Set(ByVal pd_genend As DateTime)
                l035_genend = pd_genend
            End Set
        End Property
        Public Property repfile() As String
            Get
                repfile = l035_repfile
            End Get
            Set(ByVal pc_repfile As String)
                l035_repfile = pc_repfile
            End Set
        End Property
        Public Property repdir() As String
            Get
                repdir = l035_repdir
            End Get
            Set(ByVal pc_repdir As String)
                l035_repdir = pc_repdir
            End Set
        End Property

        Public Property filesize() As Long
            Get
                filesize = l035_filesize
            End Get
            Set(ByVal pi_filesize As Long)
                l035_filesize = pi_filesize
            End Set
        End Property

        Public Property timetaken() As Long
            Get
                timetaken = l035_timetaken
            End Get
            Set(ByVal pi_timetaken As Long)
                l035_timetaken = pi_timetaken
            End Set
        End Property
        Public Property trycount() As Integer
            Get
                trycount = l035_trycount
            End Get
            Set(ByVal pi_trycount As Integer)
                l035_trycount = pi_trycount
            End Set
        End Property

        Public Property errormsg() As String
            Get
                errormsg = l035_errormsg
            End Get
            Set(ByVal pc_errormsg As String)
                l035_errormsg = pc_errormsg
            End Set
        End Property

        Public Property adc() As String
            Get
                Return l035_adc
            End Get
            Set(ByVal pc_adc As String)
                l035_adc = pc_adc
            End Set
        End Property

        Public Property repno() As String
            Get
                Return l035_repno
            End Get
            Set(ByVal pc_repno As String)
                l035_repno = pc_repno
            End Set
        End Property

        Public Property runwith() As String
            Get
                Return l035_runwith
            End Get
            Set(ByVal pc_runwith As String)
                l035_runwith = pc_runwith
            End Set
        End Property

        Public Property lds() As DataSet
            Get
                Return l035_lds
            End Get
            Set(ByVal pds As DataSet)
                l035_lds = pds
            End Set
        End Property

        Public ReadOnly Property paramlist() As List(Of param)
            Get
                paramlist = l035_paramlist
            End Get
        End Property

    End Class

    Public Class param
        Private l035_paranno As Short
        Private l035_desc As String
        Private l035_value As Object

        Public Sub New()
            l035_paranno = 0
            l035_desc = ""
            l035_value = Nothing
        End Sub

        Public Property value() As Object
            Get
                Return l035_value
            End Get
            Set(ByVal po_value As Object)
                l035_value = po_value
            End Set
        End Property
        Public Property desc() As String
            Get
                Return l035_desc
            End Get
            Set(ByVal pc_desc As String)
                l035_desc = pc_desc
            End Set
        End Property
        Public Property paramno() As Short
            Get
                Return l035_paranno
            End Get
            Set(ByVal pi_paramno As Short)
                l035_paranno = pi_paramno
            End Set
        End Property

    End Class

End Namespace



Option Strict On
Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient

Namespace CrudScaffolding
    Public Class ColumnDefinition
        Public Property Name As String
        Public Property SqlType As String
        Public Property IsPrimaryKey As Boolean
        Public Property IsIdentity As Boolean

        Public ReadOnly Property ParameterName As String
            Get
                Return "@" & Name
            End Get
        End Property
    End Class

    Public Class TableDefinition
        Public Property Name As String
        Public Property Columns As List(Of ColumnDefinition)

        Public Function PrimaryKey() As ColumnDefinition
            Return Columns.Find(Function(c) c.IsPrimaryKey)
        End Function
    End Class

    Public Class PageSettings
        Public Property MasterPageFile As String = "~/principal.master"
        Public Property ContentPlaceHolderHead As String = "head"
        Public Property ContentPlaceHolderBody As String = "ContentPlaceHolder1"
        Public Property ScriptVirtualPath As String = "js/"

        Public ReadOnly Property ListPagePrefix As String
            Get
                Return "lst_"
            End Get
        End Property

        Public ReadOnly Property FormPagePrefix As String
            Get
                Return "frm_"
            End Get
        End Property
    End Class

    Public Module CrudGenerator
        Public Property ConnectionString As String = "Data Source=(local);Initial Catalog=TuBase;Integrated Security=True;"

        Private ReadOnly SqlTypeMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"int", "Integer"},
            {"bigint", "Long"},
            {"smallint", "Short"},
            {"tinyint", "Byte"},
            {"bit", "Boolean"},
            {"decimal", "Decimal"},
            {"numeric", "Decimal"},
            {"money", "Decimal"},
            {"smallmoney", "Decimal"},
            {"float", "Double"},
            {"real", "Single"},
            {"datetime", "DateTime"},
            {"smalldatetime", "DateTime"},
            {"date", "DateTime"},
            {"datetime2", "DateTime"},
            {"uniqueidentifier", "Guid"},
            {"nvarchar", "String"},
            {"varchar", "String"},
            {"nchar", "String"},
            {"char", "String"},
            {"text", "String"},
            {"ntext", "String"}
        }

        Private ReadOnly SqlDbTypeMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"int", "SqlDbType.Int"},
            {"bigint", "SqlDbType.BigInt"},
            {"smallint", "SqlDbType.SmallInt"},
            {"tinyint", "SqlDbType.TinyInt"},
            {"bit", "SqlDbType.Bit"},
            {"decimal", "SqlDbType.Decimal"},
            {"numeric", "SqlDbType.Decimal"},
            {"money", "SqlDbType.Money"},
            {"smallmoney", "SqlDbType.SmallMoney"},
            {"float", "SqlDbType.Float"},
            {"real", "SqlDbType.Real"},
            {"datetime", "SqlDbType.DateTime"},
            {"smalldatetime", "SqlDbType.SmallDateTime"},
            {"date", "SqlDbType.Date"},
            {"datetime2", "SqlDbType.DateTime2"},
            {"uniqueidentifier", "SqlDbType.UniqueIdentifier"},
            {"nvarchar", "SqlDbType.NVarChar"},
            {"varchar", "SqlDbType.VarChar"},
            {"nchar", "SqlDbType.NChar"},
            {"char", "SqlDbType.Char"},
            {"text", "SqlDbType.Text"},
            {"ntext", "SqlDbType.NText"}
        }

        Public Function GenerateEntityClass(table As TableDefinition, namespaceName As String) As String
            Dim pk As ColumnDefinition = table.PrimaryKey()
            Dim dataClassName As String = "cls_" & table.Name
            Dim builder As New StringBuilder()

            builder.AppendLine("Option Strict On")
            builder.AppendLine("Imports System")
            builder.AppendLine("Imports System.Data")
            builder.AppendLine("Imports System.Data.SqlClient")
            builder.AppendLine()
            builder.AppendLine("Namespace " & namespaceName)
            builder.AppendLine("    Public Class " & dataClassName)
            builder.AppendLine("        ' Clase generada automáticamente para CRUD básico de " & table.Name)
            builder.AppendLine()
            builder.AppendLine("#Region \"VARIABLES\"")
            builder.AppendLine("        Private ReadOnly var_nombre_Tabla As String = \"" & table.Name & "\"")
            builder.AppendLine("        Private ReadOnly var_Campo_Id As String = \"" & pk.Name & "\"")
            builder.AppendLine("        Private ReadOnly var_Campos As String = \"" & BuildFieldList(table.Columns) & "\"")
            builder.AppendLine("        Private ReadOnly var_ConnectionString As String = \"" & ConnectionString.Replace("\"", "\"\"") & "\"")
            builder.AppendLine()

            For Each column As ColumnDefinition In table.Columns
                builder.AppendLine("        Private var_" & column.Name & " As " & MapToDotNetType(column.SqlType) & " = " & DefaultValueExpression(column))
            Next

            builder.AppendLine("#End Region")
            builder.AppendLine()
            builder.AppendLine("#Region \"PROPIEDADES\"")
            builder.AppendLine("        Public Shared ReadOnly Property Campo_Id() As String")
            builder.AppendLine("            Get")
            builder.AppendLine("                Return \"" & pk.Name & "\"")
            builder.AppendLine("            End Get")
            builder.AppendLine("        End Property")
            builder.AppendLine()

            builder.AppendLine("        Public ReadOnly Property " & pk.Name & "() As " & MapToDotNetType(pk.SqlType))
            builder.AppendLine("            Get")
            builder.AppendLine("                Return Me.var_" & pk.Name)
            builder.AppendLine("            End Get")
            builder.AppendLine("        End Property")
            builder.AppendLine()

            For Each column As ColumnDefinition In table.Columns
                If column.IsPrimaryKey Then
                    Continue For
                End If
                builder.AppendLine("        Public Property " & column.Name & "() As " & MapToDotNetType(column.SqlType))
                builder.AppendLine("            Get")
                builder.AppendLine("                Return Me.var_" & column.Name)
                builder.AppendLine("            End Get")
                builder.AppendLine("            Set(ByVal value As " & MapToDotNetType(column.SqlType) & ")")
                builder.AppendLine("                Me.var_" & column.Name & " = value")
                builder.AppendLine("            End Set")
                builder.AppendLine("        End Property")
                builder.AppendLine()
            Next
            builder.AppendLine("#End Region")
            builder.AppendLine()
            builder.AppendLine("#Region \"FUNCIONES\"")
            builder.AppendLine("        Public Sub New(Optional ByVal value As " & MapToDotNetType(pk.SqlType) & " = " & DefaultValueExpression(pk) & ")")
            builder.AppendLine("            If Not value.Equals(" & DefaultValueExpression(pk) & ") Then")
            builder.AppendLine("                Cargar(value)")
            builder.AppendLine("            End If")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Public Sub Cargar(ByVal value As " & MapToDotNetType(pk.SqlType) & ")")
            builder.AppendLine("            Using conn As New SqlConnection(var_ConnectionString)")
            builder.AppendLine("                Using cmd As New SqlCommand(\"SELECT " & BuildColumnList(table.Columns) & " FROM " & table.Name & " WHERE " & pk.Name & " = @" & pk.Name & "\", conn)")
            builder.AppendLine("                    cmd.Parameters.AddWithValue(\"@" & pk.Name & "\", value)")
            builder.AppendLine("                    conn.Open()")
            builder.AppendLine("                    Using reader As SqlDataReader = cmd.ExecuteReader()")
            builder.AppendLine("                        If reader.Read() Then")
            builder.AppendLine("                            MapFromReader(reader)")
            builder.AppendLine("                        Else")
            builder.AppendLine("                            Me.var_" & pk.Name & " = " & DefaultValueExpression(pk))
            builder.AppendLine("                        End If")
            builder.AppendLine("                    End Using")
            builder.AppendLine("                End Using")
            builder.AppendLine("            End Using")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Private Sub MapFromReader(reader As SqlDataReader)")
            For Each column As ColumnDefinition In table.Columns
                builder.AppendLine("            Me.var_" & column.Name & " = " & ReadValueExpression(column))
            Next
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Public Function Guardar(ByRef mensaje As String) As Boolean")
            builder.AppendLine("            Try")
            builder.AppendLine("                If Me.var_" & pk.Name & ".Equals(" & DefaultValueExpression(pk) & ") Then")
            builder.AppendLine("                    Return Insertar(mensaje)")
            builder.AppendLine("                Else")
            builder.AppendLine("                    Return Actualizar(mensaje)")
            builder.AppendLine("                End If")
            builder.AppendLine("            Catch ex As Exception")
            builder.AppendLine("                mensaje = ex.Message")
            builder.AppendLine("                Return False")
            builder.AppendLine("            End Try")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function Insertar(ByRef mensaje As String) As Boolean")
            builder.AppendLine("            Using conn As New SqlConnection(var_ConnectionString)")
            builder.AppendLine("                Dim sql As String = \"" & BuildInsertStatementWithOutput(table) & "\"")
            builder.AppendLine("                Using cmd As New SqlCommand(sql, conn)")
            For Each column As ColumnDefinition In table.Columns
                If column.IsIdentity Then
                    Continue For
                End If
                builder.AppendLine("                    cmd.Parameters.Add(\"" & column.ParameterName & "\", " & MapToSqlDbType(column.SqlType) & ").Value = Me.var_" & column.Name)
            Next
            builder.AppendLine("                    conn.Open()")
            builder.AppendLine("                    Dim result As Object = cmd.ExecuteScalar()")
            builder.AppendLine("                    If result IsNot Nothing AndAlso Not Convert.IsDBNull(result) Then")
            builder.AppendLine("                        Me.var_" & pk.Name & " = CType(result, " & MapToDotNetType(pk.SqlType) & ")")
            builder.AppendLine("                    End If")
            builder.AppendLine("                    Return True")
            builder.AppendLine("                End Using")
            builder.AppendLine("            End Using")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function Actualizar(ByRef mensaje As String) As Boolean")
            builder.AppendLine("            Using conn As New SqlConnection(var_ConnectionString)")
            builder.AppendLine("                Dim sql As String = \"" & BuildUpdateStatement(table) & "\"")
            builder.AppendLine("                Using cmd As New SqlCommand(sql, conn)")
            For Each column As ColumnDefinition In table.Columns
                If column.IsPrimaryKey Then
                    Continue For
                End If
                If column.IsIdentity Then
                    Continue For
                End If
                builder.AppendLine("                    cmd.Parameters.Add(\"" & column.ParameterName & "\", " & MapToSqlDbType(column.SqlType) & ").Value = Me.var_" & column.Name)
            Next
            builder.AppendLine("                    cmd.Parameters.Add(\"@" & pk.Name & "\", " & MapToSqlDbType(pk.SqlType) & ").Value = Me.var_" & pk.Name)
            builder.AppendLine("                    conn.Open()")
            builder.AppendLine("                    cmd.ExecuteNonQuery()")
            builder.AppendLine("                    Return True")
            builder.AppendLine("                End Using")
            builder.AppendLine("            End Using")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Public Shared Function Eliminar(ByVal value As " & MapToDotNetType(pk.SqlType) & ", ByRef mensaje As String) As Boolean")
            builder.AppendLine("            Try")
            builder.AppendLine("                Using conn As New SqlConnection(\"" & ConnectionString.Replace("\"", "\"\"") & "\")")
            builder.AppendLine("                    Using cmd As New SqlCommand(\"DELETE FROM " & table.Name & " WHERE " & pk.Name & " = @" & pk.Name & "\", conn)")
            builder.AppendLine("                        cmd.Parameters.AddWithValue(\"@" & pk.Name & "\", value)")
            builder.AppendLine("                        conn.Open()")
            builder.AppendLine("                        cmd.ExecuteNonQuery()")
            builder.AppendLine("                        Return True")
            builder.AppendLine("                    End Using")
            builder.AppendLine("                End Using")
            builder.AppendLine("            Catch ex As Exception")
            builder.AppendLine("                mensaje = ex.Message")
            builder.AppendLine("                Return False")
            builder.AppendLine("            End Try")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Public Shared Function Listar(ByRef mensaje As String) As DataTable")
            builder.AppendLine("            Dim dt As New DataTable()")
            builder.AppendLine("            Try")
            builder.AppendLine("                Using conn As New SqlConnection(\"" & ConnectionString.Replace("\"", "\"\"") & "\")")
            builder.AppendLine("                    Using cmd As New SqlCommand(\"SELECT " & BuildColumnList(table.Columns) & " FROM " & table.Name & "\", conn)")
            builder.AppendLine("                        Using da As New SqlDataAdapter(cmd)")
            builder.AppendLine("                            da.Fill(dt)")
            builder.AppendLine("                        End Using")
            builder.AppendLine("                    End Using")
            builder.AppendLine("                End Using")
            builder.AppendLine("            Catch ex As Exception")
            builder.AppendLine("                mensaje = ex.Message")
            builder.AppendLine("            End Try")
            builder.AppendLine("            Return dt")
            builder.AppendLine("        End Function")
            builder.AppendLine("#End Region")
            builder.AppendLine("    End Class")
            builder.AppendLine("End Namespace")

            Return builder.ToString()
        End Function

        Public Function GenerateAspxMarkup(table As TableDefinition, namespaceName As String, Optional pageSettings As PageSettings = Nothing) As String
            Dim settings As PageSettings = If(pageSettings, New PageSettings())
            Dim builder As New StringBuilder()
            builder.AppendLine("<%@ Page Language=\"VB\" MasterPageFile=\"" & settings.MasterPageFile & "\" AutoEventWireup=\"false\" CodeFile=\"" & settings.FormPagePrefix & table.Name & ".aspx.vb\" Inherits=\"" & namespaceName & "." & table.Name & "Page\" %>")
            builder.AppendLine()
            builder.AppendLine("<asp:Content ID=\"Content1\" ContentPlaceHolderID=\"" & settings.ContentPlaceHolderHead & "\" runat=\"Server\">")
            builder.AppendLine("    <style type=\"text/css\" title=\"currentStyle\">")
            builder.AppendLine("        @import \"css/jquery.dataTables.css\";")
            builder.AppendLine("    </style>")
            builder.AppendLine("</asp:Content>")
            builder.AppendLine()
            builder.AppendLine("<asp:Content ID=\"Content2\" ContentPlaceHolderID=\"" & settings.ContentPlaceHolderBody & "\" runat=\"Server\">")
            builder.AppendLine("    <div class=\"container\">")
            builder.AppendLine("        <h2>Administración de " & table.Name & "</h2>")
            builder.AppendLine("        <asp:Label ID=\"lblStatus\" runat=\"server\" EnableViewState=\"false\" ForeColor=\"Red\"></asp:Label>")
            builder.AppendLine("        <asp:Panel ID=\"pnlForm\" runat=\"server\">")
            For Each column As ColumnDefinition In table.Columns
                If column.IsIdentity Then
                    Continue For
                End If
                builder.AppendLine("            <div class=\"control-group\">")
                builder.AppendLine("                <asp:Label ID=\"lbl" & column.Name & "\" CssClass=\"control-label\" runat=\"server\" AssociatedControlID=\"txt" & column.Name & "\" Text=\"" & column.Name & ":\"></asp:Label>")
                builder.AppendLine("                <div class=\"controls\">")
                builder.AppendLine("                    <asp:TextBox ID=\"txt" & column.Name & "\" CssClass=\"form-control\" runat=\"server\"></asp:TextBox>")
                builder.AppendLine("                </div>")
                builder.AppendLine("            </div>")
            Next
            builder.AppendLine("            <asp:Button ID=\"btnSave\" runat=\"server\" Text=\"Guardar\" CssClass=\"btn btn-primary\" OnClick=\"btnSave_Click\" />")
            builder.AppendLine("            <asp:Button ID=\"btnClear\" runat=\"server\" Text=\"Limpiar\" CssClass=\"btn btn-default\" OnClick=\"btnClear_Click\" />")
            builder.AppendLine("        </asp:Panel>")
            builder.AppendLine("        <hr />")
            builder.AppendLine("        <asp:GridView ID=\"GridView1\" runat=\"server\" AutoGenerateColumns=\"False\" DataKeyNames=\"" & table.PrimaryKey().Name & "\" OnRowEditing=\"GridView1_RowEditing\" OnRowCancelingEdit=\"GridView1_RowCancelingEdit\" OnRowUpdating=\"GridView1_RowUpdating\" OnRowDeleting=\"GridView1_RowDeleting\" CellPadding=\"4\" GridLines=\"None\">")
            builder.AppendLine("            <Columns>")
            For Each column As ColumnDefinition In table.Columns
                builder.AppendLine("                <asp:BoundField DataField=\"" & column.Name & "\" HeaderText=\"" & column.Name & "\" ReadOnly=\"" & column.IsIdentity.ToString().ToLowerInvariant() & "\" />")
            Next
            builder.AppendLine("                <asp:CommandField ShowEditButton=\"True\" ShowDeleteButton=\"True\" />")
            builder.AppendLine("            </Columns>")
            builder.AppendLine("            <AlternatingRowStyle BackColor=\"#EFEFEF\" />")
            builder.AppendLine("            <HeaderStyle BackColor=\"#336699\" ForeColor=\"White\" Font-Bold=\"True\" />")
            builder.AppendLine("        </asp:GridView>")
            builder.AppendLine("    </div>")
            builder.AppendLine("    <script src=\"<%= CacheHelper.GetVersionedUrl(\"" & settings.ScriptVirtualPath & settings.FormPagePrefix & table.Name & ".js\", HttpContext.Current) %>\"></script>")
            builder.AppendLine("</asp:Content>")
            Return builder.ToString()
        End Function

        Public Function GenerateCodeBehind(table As TableDefinition, namespaceName As String) As String
            Dim pk As ColumnDefinition = table.PrimaryKey()
            Dim dataClassName As String = "cls_" & table.Name
            Dim builder As New StringBuilder()
            builder.AppendLine("Option Strict On")
            builder.AppendLine("Imports System")
            builder.AppendLine("Imports System.Data")
            builder.AppendLine("Imports System.Data.SqlClient")
            builder.AppendLine()
            builder.AppendLine("Namespace " & namespaceName)
            builder.AppendLine("    Public Class " & table.Name & "Page")
            builder.AppendLine("        Inherits System.Web.UI.Page")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load")
            builder.AppendLine("            If Not IsPostBack Then")
            builder.AppendLine("                BindGrid()")
            builder.AppendLine("            End If")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Private Sub BindGrid()")
            builder.AppendLine("            Dim mensaje As String = String.Empty")
            builder.AppendLine("            GridView1.DataSource = " & dataClassName & ".Listar(mensaje)")
            builder.AppendLine("            GridView1.DataBind()")
            builder.AppendLine("            ShowMessage(mensaje)")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs)")
            builder.AppendLine("            Dim mensaje As String = String.Empty")
            builder.AppendLine("            Dim registro As New " & dataClassName & "()")

            For Each column As ColumnDefinition In table.Columns
                If column.IsIdentity Then
                    Continue For
                End If
                builder.AppendLine("            " & BuildAssignmentFromForm(column, "registro"))
            Next

            builder.AppendLine("            If registro.Guardar(mensaje) Then")
            builder.AppendLine("                ClearForm()")
            builder.AppendLine("                BindGrid()")
            builder.AppendLine("            End If")
            builder.AppendLine("            ShowMessage(mensaje)")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As EventArgs)")
            builder.AppendLine("            ClearForm()")
            builder.AppendLine("            ShowMessage(String.Empty)")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Private Sub ClearForm()")

            For Each column As ColumnDefinition In table.Columns
                If column.IsIdentity Then
                    Continue For
                End If
                builder.AppendLine("            txt" & column.Name & ".Text = String.Empty")
            Next

            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs)")
            builder.AppendLine("            GridView1.EditIndex = e.NewEditIndex")
            builder.AppendLine("            BindGrid()")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs)")
            builder.AppendLine("            GridView1.EditIndex = -1")
            builder.AppendLine("            BindGrid()")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs)")
            builder.AppendLine("            Dim mensaje As String = String.Empty")
            builder.AppendLine("            Dim key As Object = GridView1.DataKeys(e.RowIndex).Value")
            builder.AppendLine("            Dim row As System.Web.UI.WebControls.GridViewRow = GridView1.Rows(e.RowIndex)")
            builder.AppendLine("            Dim registro As New " & dataClassName & "()")
            builder.AppendLine("            registro.Cargar(CType(key, " & MapToDotNetType(pk.SqlType) & "))")

            Dim columnIndex As Integer = 0
            For Each column As ColumnDefinition In table.Columns
                If column.IsPrimaryKey Then
                    columnIndex += 1
                    Continue For
                End If
                If column.IsIdentity Then
                    columnIndex += 1
                    Continue For
                End If
                builder.AppendLine("            " & BuildAssignmentFromGrid(column, "registro", columnIndex))
                columnIndex += 1
            Next

            builder.AppendLine("            If registro.Guardar(mensaje) Then")
            builder.AppendLine("                GridView1.EditIndex = -1")
            builder.AppendLine("                BindGrid()")
            builder.AppendLine("            End If")
            builder.AppendLine("            ShowMessage(mensaje)")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)")
            builder.AppendLine("            Dim mensaje As String = String.Empty")
            builder.AppendLine("            Dim key As Object = GridView1.DataKeys(e.RowIndex).Value")
            builder.AppendLine("            If " & dataClassName & ".Eliminar(CType(key, " & MapToDotNetType(pk.SqlType) & "), mensaje) Then")
            builder.AppendLine("                BindGrid()")
            builder.AppendLine("            End If")
            builder.AppendLine("            ShowMessage(mensaje)")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Private Sub ShowMessage(message As String)")
            builder.AppendLine("            lblStatus.Text = message")
            builder.AppendLine("        End Sub")
            builder.AppendLine()
            builder.AppendLine("        Private Function BuildText(ByVal control As System.Web.UI.WebControls.TextBox) As String")
            builder.AppendLine("            If control Is Nothing Then")
            builder.AppendLine("                Return String.Empty")
            builder.AppendLine("            End If")
            builder.AppendLine("            Return control.Text.Trim()")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToInteger(ByVal value As String) As Integer")
            builder.AppendLine("            Dim result As Integer = 0")
            builder.AppendLine("            Integer.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToLongValue(ByVal value As String) As Long")
            builder.AppendLine("            Dim result As Long = 0")
            builder.AppendLine("            Long.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToShortValue(ByVal value As String) As Short")
            builder.AppendLine("            Dim result As Short = 0")
            builder.AppendLine("            Short.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToByteValue(ByVal value As String) As Byte")
            builder.AppendLine("            Dim result As Byte = 0")
            builder.AppendLine("            Byte.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToDecimalValue(ByVal value As String) As Decimal")
            builder.AppendLine("            Dim result As Decimal = 0")
            builder.AppendLine("            Decimal.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToDoubleValue(ByVal value As String) As Double")
            builder.AppendLine("            Dim result As Double = 0")
            builder.AppendLine("            Double.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToSingleValue(ByVal value As String) As Single")
            builder.AppendLine("            Dim result As Single = 0")
            builder.AppendLine("            Single.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToBooleanValue(ByVal value As String) As Boolean")
            builder.AppendLine("            Dim result As Boolean = False")
            builder.AppendLine("            Boolean.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToDateValue(ByVal value As String) As DateTime")
            builder.AppendLine("            Dim result As DateTime = DateTime.MinValue")
            builder.AppendLine("            DateTime.TryParse(value, result)")
            builder.AppendLine("            Return result")
            builder.AppendLine("        End Function")
            builder.AppendLine()
            builder.AppendLine("        Private Function ToGuidValue(ByVal value As String) As Guid")
            builder.AppendLine("            Dim result As Guid = Guid.Empty")
            builder.AppendLine("            If Guid.TryParse(value, result) Then")
            builder.AppendLine("                Return result")
            builder.AppendLine("            End If")
            builder.AppendLine("            Return Guid.Empty")
            builder.AppendLine("        End Function")
            builder.AppendLine("    End Class")
            builder.AppendLine("End Namespace")
            Return builder.ToString()
        End Function

        Public Sub WriteArtifacts(table As TableDefinition, namespaceName As String, outputRoot As String, Optional pageSettings As PageSettings = Nothing)
            Dim targetDir As String = Path.Combine(outputRoot, table.Name)
            Dim settings As PageSettings = If(pageSettings, New PageSettings())
            If Not Directory.Exists(targetDir) Then
                Directory.CreateDirectory(targetDir)
            End If

            File.WriteAllText(Path.Combine(targetDir, table.Name & ".vb"), GenerateEntityClass(table, namespaceName), Encoding.UTF8)
            File.WriteAllText(Path.Combine(targetDir, settings.FormPagePrefix & table.Name & ".aspx"), GenerateAspxMarkup(table, namespaceName, settings), Encoding.UTF8)
            File.WriteAllText(Path.Combine(targetDir, settings.FormPagePrefix & table.Name & ".aspx.vb"), GenerateCodeBehind(table, namespaceName), Encoding.UTF8)
        End Sub

        Private Function MapToDotNetType(sqlType As String) As String
            Dim key As String = CleanSqlType(sqlType)
            If SqlTypeMap.ContainsKey(key) Then
                Return SqlTypeMap(key)
            End If
            Return "String"
        End Function

        Private Function MapToSqlDbType(sqlType As String) As String
            Dim key As String = CleanSqlType(sqlType)
            If SqlDbTypeMap.ContainsKey(key) Then
                Return SqlDbTypeMap(key)
            End If
            Return "SqlDbType.Variant"
        End Function

        Private Function CleanSqlType(sqlType As String) As String
            Dim value As String = sqlType
            Dim parenIndex As Integer = value.IndexOf("("c)
            If parenIndex > 0 Then
                value = value.Substring(0, parenIndex)
            End If
            Return value.Trim().ToLowerInvariant()
        End Function

        Private Function BuildColumnList(columns As List(Of ColumnDefinition)) As String
            Dim builder As New StringBuilder()
            For i As Integer = 0 To columns.Count - 1
                builder.Append(columns(i).Name)
                If i < columns.Count - 1 Then
                    builder.Append(", ")
                End If
            Next
            Return builder.ToString()
        End Function

        Private Function BuildFieldList(columns As List(Of ColumnDefinition)) As String
            Dim names As New List(Of String)()
            For Each column As ColumnDefinition In columns
                names.Add(column.Name)
            Next
            Return String.Join(",", names.ToArray())
        End Function

        Private Function BuildInsertStatementWithOutput(table As TableDefinition) As String
            Dim cols As New List(Of String)()
            Dim values As New List(Of String)()
            For Each column As ColumnDefinition In table.Columns
                If column.IsIdentity Then
                    Continue For
                End If
                cols.Add(column.Name)
                values.Add(column.ParameterName)
            Next
            Dim pk As ColumnDefinition = table.PrimaryKey()
            Dim sb As New StringBuilder()
            sb.Append("INSERT INTO " & table.Name & " (")
            sb.Append(String.Join(", ", cols.ToArray()))
            sb.Append(") OUTPUT INSERTED." & pk.Name & " VALUES (")
            sb.Append(String.Join(", ", values.ToArray()))
            sb.Append(")")
            Return sb.ToString()
        End Function

        Private Function BuildUpdateStatement(table As TableDefinition) As String
            Dim sets As New List(Of String)()
            Dim pk As ColumnDefinition = table.PrimaryKey()

            For Each column As ColumnDefinition In table.Columns
                If column.IsPrimaryKey OrElse column.IsIdentity Then
                    Continue For
                End If
                sets.Add(column.Name & " = " & column.ParameterName)
            Next

            Dim sb As New StringBuilder()
            sb.Append("UPDATE " & table.Name & " SET ")
            sb.Append(String.Join(", ", sets.ToArray()))
            sb.Append(" WHERE " & pk.Name & " = " & pk.ParameterName)
            Return sb.ToString()
        End Function

        Private Function DefaultValueExpression(column As ColumnDefinition) As String
            Dim typeName As String = MapToDotNetType(column.SqlType)
            Select Case typeName
                Case "Integer"
                    Return "0"
                Case "Long"
                    Return "0"
                Case "Short"
                    Return "0"
                Case "Byte"
                    Return "0"
                Case "Decimal"
                    Return "0D"
                Case "Double"
                    Return "0R"
                Case "Single"
                    Return "0F"
                Case "Boolean"
                    Return "False"
                Case "DateTime"
                    Return "Date.MinValue"
                Case "Guid"
                    Return "Guid.Empty"
                Case Else
                    Return "\"\""
            End Select
        End Function

        Private Function ReadValueExpression(column As ColumnDefinition) As String
            Dim typeName As String = MapToDotNetType(column.SqlType)
            Dim name As String = column.Name
            Dim defaultVal As String = DefaultValueExpression(column)

            Select Case typeName
                Case "Integer"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToInt32(reader(\"" & name & "\")))"
                Case "Long"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToInt64(reader(\"" & name & "\")))"
                Case "Short"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToInt16(reader(\"" & name & "\")))"
                Case "Byte"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToByte(reader(\"" & name & "\")))"
                Case "Decimal"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToDecimal(reader(\"" & name & "\")))"
                Case "Double"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToDouble(reader(\"" & name & "\")))"
                Case "Single"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToSingle(reader(\"" & name & "\")))"
                Case "Boolean"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToBoolean(reader(\"" & name & "\")))"
                Case "DateTime"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", Convert.ToDateTime(reader(\"" & name & "\")))"
                Case "Guid"
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), Guid.Empty, Guid.Parse(reader(\"" & name & "\").ToString()))"
                Case Else
                    Return "If(reader.IsDBNull(reader.GetOrdinal(\"" & name & "\")), " & defaultVal & ", reader(\"" & name & "\").ToString())"
            End Select
        End Function

        Private Function BuildAssignmentFromForm(column As ColumnDefinition, instanceName As String) As String
            Return instanceName & "." & column.Name & " = " & BuildValueFromForm(column)
        End Function

        Private Function BuildAssignmentFromGrid(column As ColumnDefinition, instanceName As String, columnIndex As Integer) As String
            Return instanceName & "." & column.Name & " = " & BuildValueFromGrid(column, columnIndex)
        End Function

        Private Function BuildValueFromForm(column As ColumnDefinition) As String
            Dim controlName As String = "txt" & column.Name
            Dim typeName As String = MapToDotNetType(column.SqlType)

            Select Case typeName
                Case "Integer"
                    Return "ToInteger(BuildText(" & controlName & "))"
                Case "Long"
                    Return "ToLongValue(BuildText(" & controlName & "))"
                Case "Short"
                    Return "ToShortValue(BuildText(" & controlName & "))"
                Case "Byte"
                    Return "ToByteValue(BuildText(" & controlName & "))"
                Case "Decimal"
                    Return "ToDecimalValue(BuildText(" & controlName & "))"
                Case "Double"
                    Return "ToDoubleValue(BuildText(" & controlName & "))"
                Case "Single"
                    Return "ToSingleValue(BuildText(" & controlName & "))"
                Case "Boolean"
                    Return "ToBooleanValue(BuildText(" & controlName & "))"
                Case "DateTime"
                    Return "ToDateValue(BuildText(" & controlName & "))"
                Case "Guid"
                    Return "ToGuidValue(BuildText(" & controlName & "))"
                Case Else
                    Return "BuildText(" & controlName & ")"
            End Select
        End Function

        Private Function BuildValueFromGrid(column As ColumnDefinition, columnIndex As Integer) As String
            Dim baseExpression As String = "BuildText(CType(row.Cells(" & columnIndex.ToString() & ").Controls(0), System.Web.UI.WebControls.TextBox))"
            Dim typeName As String = MapToDotNetType(column.SqlType)

            Select Case typeName
                Case "Integer"
                    Return "ToInteger(" & baseExpression & ")"
                Case "Long"
                    Return "ToLongValue(" & baseExpression & ")"
                Case "Short"
                    Return "ToShortValue(" & baseExpression & ")"
                Case "Byte"
                    Return "ToByteValue(" & baseExpression & ")"
                Case "Decimal"
                    Return "ToDecimalValue(" & baseExpression & ")"
                Case "Double"
                    Return "ToDoubleValue(" & baseExpression & ")"
                Case "Single"
                    Return "ToSingleValue(" & baseExpression & ")"
                Case "Boolean"
                    Return "ToBooleanValue(" & baseExpression & ")"
                Case "DateTime"
                    Return "ToDateValue(" & baseExpression & ")"
                Case "Guid"
                    Return "ToGuidValue(" & baseExpression & ")"
                Case Else
                    Return baseExpression
            End Select
        End Function
    End Module
End Namespace

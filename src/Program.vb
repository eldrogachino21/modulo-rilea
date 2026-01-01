Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports CrudScaffolding

Module Program
    Sub Main()
        Dim transferTable As New TableDefinition() With {
            .Name = "tbl_traspasos",
            .Columns = New List(Of ColumnDefinition)() From {
                New ColumnDefinition() With {.Name = "id", .SqlType = "int", .IsPrimaryKey = True, .IsIdentity = True},
                New ColumnDefinition() With {.Name = "id_centro", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_membresia", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_alumno", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "hrs_total", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "hrs_usadas", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "hrs_restantes", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "motivo_traspaso", .SqlType = "nvarchar(200)"},
                New ColumnDefinition() With {.Name = "nombre", .SqlType = "nvarchar(120)"},
                New ColumnDefinition() With {.Name = "apellido", .SqlType = "nvarchar(120)"},
                New ColumnDefinition() With {.Name = "rfc", .SqlType = "varchar(20)"},
                New ColumnDefinition() With {.Name = "fecha_nacimiento", .SqlType = "date"},
                New ColumnDefinition() With {.Name = "correo", .SqlType = "varchar(150)"},
                New ColumnDefinition() With {.Name = "telefono", .SqlType = "varchar(60)"},
                New ColumnDefinition() With {.Name = "sexo", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_capacitacion", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_curso", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_grupo", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "estatus", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "id_usuario_autorizacion", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "fecha_autorizacion", .SqlType = "datetime"},
                New ColumnDefinition() With {.Name = "motivo_rechazo", .SqlType = "nvarchar(200)"},
                New ColumnDefinition() With {.Name = "id_usuario_registro", .SqlType = "int"},
                New ColumnDefinition() With {.Name = "numero_membresia", .SqlType = "varchar(120)"},
                New ColumnDefinition() With {.Name = "direccion", .SqlType = "nvarchar(200)"},
                New ColumnDefinition() With {.Name = "curp", .SqlType = "varchar(20)"},
                New ColumnDefinition() With {.Name = "nacionalidad", .SqlType = "int"}
            }
        }

        Dim namespaceName As String = "DemoApp.Web"
        CrudGenerator.ConnectionString = "Data Source=(local);Initial Catalog=TuBase;Integrated Security=True;"

        Dim outputRoot As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output")
        CrudGenerator.WriteArtifacts(transferTable, namespaceName, outputRoot)

        Console.WriteLine("CRUD generado para la tabla '" & transferTable.Name & "'.")
        Console.WriteLine("Revisa: " & Path.Combine(outputRoot, transferTable.Name))
    End Sub
End Module

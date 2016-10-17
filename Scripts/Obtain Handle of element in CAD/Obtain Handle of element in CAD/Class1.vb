
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Interop

Public Class Class1

    Public AcadApp As Autodesk.AutoCAD.Interop.AcadApplication
    Public AcadDoc As Autodesk.AutoCAD.Interop.AcadDocument


    <Autodesk.AutoCAD.Runtime.CommandMethod("HandleID")> Public Sub Handle1()

        AcadApp = GetObject(, "Autocad.Application")
        AcadDoc = AcadApp.ActiveDocument

        Dim Form2 As Form1
        Form2 = New Form1

        Dim entHandle As AcadEntity
        Dim ss As AcadSelectionSet
        Dim nameent As String


        ss = AcadDoc.ActiveSelectionSet

        entHandle = ss.Item(0)

        nameent = entHandle.Handle

        Form2.TextBox1.Text = nameent
        Form2.Show()

    End Sub



End Class

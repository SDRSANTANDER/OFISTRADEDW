Imports System

Public Class clsUtil

    Public Shared Sub LiberarObjetoCOM(ByRef oObjCOM As Object)
        'Liberar y destruir Objeto COM
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            GC.Collect()
        End If
    End Sub

End Class

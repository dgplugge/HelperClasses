' Author: Donald G Plugge
' Date: 12/12/08
Imports System.Windows.Forms

Public Class MyClipBoard

    ' dgp rev 12/12/08 Selected Cells
    Private mCells As DataGridViewSelectedCellCollection
    Public Property Cells() As DataGridViewSelectedCellCollection
        Get
            Return mCells
        End Get
        Set(ByVal value As DataGridViewSelectedCellCollection)
            mCells = value
        End Set
    End Property

    ' dgp rev 12/12/08 Upper
    Private mUpper As Int16
    Public Property Upper() As Int16
        Get
            Return mUpper
        End Get
        Set(ByVal value As Int16)
            mUpper = value
        End Set
    End Property

    ' dgp rev 12/12/08 Left
    Private mLeft As Int16
    Public Property Left() As Int16
        Get
            Return mLeft
        End Get
        Set(ByVal value As Int16)
            mLeft = value
        End Set
    End Property
    ' dgp rev 12/12/08 Rows
    Private mRowCnt As Int16
    Public Property RowCnt() As Int16
        Get
            Return mRowCnt
        End Get
        Set(ByVal value As Int16)
            mRowCnt = value
        End Set
    End Property
    ' dgp rev 12/12/08 Cols
    Private mColCnt As Int16
    Public Property ColCnt() As Int16
        Get
            Return mColCnt
        End Get
        Set(ByVal value As Int16)
            mColCnt = value
        End Set
    End Property
    ' dgp rev 12/12/08 Valid Selection
    Private mValid As Boolean = False
    Public Property Valid() As Boolean
        Get
            Return mValid
        End Get
        Set(ByVal value As Boolean)
            mValid = value
        End Set
    End Property

    ' dgp rev 12/12/08 Paste Clipboard onto Current location
    Private Sub UpperLeft()

        ' dgp rev 12/12/08 due to the nature of the selected cells
        ' we need to calculate the upper left corner if more than one cell selected
        mUpper = mCells.Item(0).RowIndex
        mLeft = mCells.Item(0).ColumnIndex
        If (mCells.Count = 1) Then Exit Sub

        Dim item

        Dim Xes As New ArrayList
        Dim Yes As New ArrayList
        For Each item In mCells
            If (Not Xes.Contains(item.ColumnIndex)) Then Xes.Add(item.ColumnIndex)
            If (Not Yes.Contains(item.RowIndex)) Then Yes.Add(item.RowIndex)
        Next
        Xes.Sort()
        Yes.Sort()
        mLeft = Xes(0)
        mUpper = Yes(0)
        mColCnt = Xes.Count
        mRowCnt = Yes.Count

    End Sub

    ' dgp rev 12/12/08 New MyClipBoard instance
    Public Sub New(ByVal mUserCells As DataGridViewSelectedCellCollection)

        mCells = mUserCells
        UpperLeft()

    End Sub
End Class

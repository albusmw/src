Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Universal class enabling file drag-and-drop for text and list boxes.</summary>
    Public Class DragDrop

        Public Event DropOccured(ByVal Files As String())

        '''<summary>TRUE: Fill the list / FALSE: Raise events only.</summary>
        Public Property FillList() As Boolean
            Get
                Return MyFillList
            End Get
            Set(value As Boolean)
                MyFillList = value
            End Set
        End Property
        Private MyFillList As Boolean = True

        '''<summary>Add the standard handling for drag-n-drop to the passed text box.</summary>
        Public Sub New(ByRef Box As Windows.Forms.TextBox)
            Me.New(Box, True)
        End Sub

        '''<summary>Add the standard handling for drag-n-drop to the passed text box.</summary>
        '''<param name="FillListToSet">TRUE: Fill the list / FALSE: Raise events only.</param>
        Public Sub New(ByRef Box As Windows.Forms.TextBox, ByVal FillListToSet As Boolean)
            AllowDrop(Box)
            AddHandler Box.DragEnter, AddressOf DragEnter_AcceptFile
            AddHandler Box.DragDrop, AddressOf DragDrop_GetFile
            Me.FillList = FillListToSet
        End Sub

        '''<summary>Add the standard handling for drag-n-drop to the passed text box.</summary>
        Public Sub New(ByRef Box As Windows.Forms.ListBox)
            AllowDrop(Box)
            AddHandler Box.DragEnter, AddressOf DragEnter_AcceptFile
            AddHandler Box.DragDrop, AddressOf DragDrop_GetFile
        End Sub

        '''<summary>Add the standard handling for drag-n-drop to the passed text box.</summary>
        Public Sub New(ByRef Box As Windows.Forms.CheckedListBox)
            AllowDrop(Box)
            AddHandler Box.DragEnter, AddressOf DragEnter_AcceptFile
            AddHandler Box.DragDrop, AddressOf DragDrop_GetFile
        End Sub

        Public Sub AllowDrop(ByRef Box As Windows.Forms.TextBox)
            Box.AllowDrop = True
        End Sub

        Public Sub AllowDrop(ByRef Box As Windows.Forms.ListBox)
            Box.AllowDrop = True
        End Sub

        Public Sub AllowDrop(ByRef Box As Windows.Forms.CheckedListBox)
            Box.AllowDrop = True
        End Sub

        '''<summary>Accept ONE file for drop - multiple files will be rejected.</summary>
        Public Sub DragEnter_AcceptFile(sender As Object, e As Windows.Forms.DragEventArgs)
            e.Effect = Windows.Forms.DragDropEffects.None
            If e.Data.GetDataPresent(Windows.Forms.DataFormats.FileDrop) Then
                If TypeOf sender Is Windows.Forms.TextBox Then
                    If CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()).Length >= 1 Then
                        e.Effect = Windows.Forms.DragDropEffects.All
                    End If
                End If
                If TypeOf sender Is Windows.Forms.ListBox Then
                    e.Effect = Windows.Forms.DragDropEffects.All
                End If
                If TypeOf sender Is Windows.Forms.CheckedListBox Then
                    e.Effect = Windows.Forms.DragDropEffects.All
                End If
            End If
        End Sub

        '''<summary>Get the (first) file that was dropped.</summary>
        Public Sub DragDrop_GetFile(sender As Object, e As Windows.Forms.DragEventArgs)
            Try
                If e.Data.GetDataPresent(Windows.Forms.DataFormats.FileDrop) Then
                    'Textbox only gets the 1st dropped file
                    If TypeOf sender Is Windows.Forms.TextBox Then
                        If FillList = True Then
                            CType(sender, Windows.Forms.TextBox).Text = CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String())(0)
                        End If
                        RaiseEvent DropOccured(CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()))
                        Exit Sub
                    End If
                    If TypeOf sender Is Windows.Forms.ListBox Then
                        If FillList = True Then
                            CType(sender, Windows.Forms.ListBox).Items.AddRange(CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()))
                        End If
                        RaiseEvent DropOccured(CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()))
                        Exit Sub
                    End If
                    If TypeOf sender Is Windows.Forms.CheckedListBox Then
                        If FillList = True Then
                            CType(sender, Windows.Forms.CheckedListBox).Items.AddRange(CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()))
                        End If
                        RaiseEvent DropOccured(CType(e.Data.GetData(Windows.Forms.DataFormats.FileDrop), String()))
                        Exit Sub
                    End If
                End If
            Catch ex As Exception
                Exit Sub
            End Try
        End Sub

    End Class

End Namespace
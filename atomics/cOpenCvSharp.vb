Option Explicit On
Option Strict On

Public Class cOpenCvSharp

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As IntPtr, ByVal pSrc As IntPtr, ByVal ByteLen As Long)

    Public Shared Sub MedianBlur(ByRef Data(,) As UInt16, ByVal kSize As Integer)

        Dim DataType As OpenCvSharp.MatType = OpenCvSharp.MatType.CV_16UC1
        Dim BytePerSample As Integer = 2

        If Data.LongLength > 0 Then
            Using src As New OpenCvSharp.Mat(Data.GetUpperBound(0) + 1, Data.GetUpperBound(1) + 1, DataType, Data)
                Dim dst As New OpenCvSharp.Mat
                OpenCvSharp.Cv2.MedianBlur(src, dst, kSize)
                Dim MyHandle As Runtime.InteropServices.GCHandle = Runtime.InteropServices.GCHandle.Alloc(Data, Runtime.InteropServices.GCHandleType.Pinned)
                CopyMemory(System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(Data, 0), dst.Data, Data.LongLength * BytePerSample)
                MyHandle.Free()
            End Using
        End If

    End Sub

End Class
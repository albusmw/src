Option Explicit On
Option Strict On

'''<summary>Use this class to dynamically call any IPP DLL.</summary>
'''<see cref=">https://software.intel.com/en-us/ipp-dev-guide"/>
'''<see cref=">https://software.intel.com/en-us/articles/descriptor-codes-and-parameters-for-ippi-functions"/>
Partial Public Class cIntelIPP

#Region "Handles and constructors"

    '''<summary>Handle to the DLL.</summary>
    Private ippsHandle As IntPtr = Nothing
    Private ippvmHandle As IntPtr = Nothing
    Private ippiHandle As IntPtr = Nothing

    '''<summary>Error message that came up during loading.</summary>
    Public ReadOnly Property LoadError() As String
        Get
            Return MyLoadError
        End Get
    End Property
    Private MyLoadError As String = String.Empty

    '''<summary>Indicate that a DLL handle could be created.</summary>
    Public ReadOnly Property DLLHandleValid() As Boolean
        Get
            If IsNothing(ippsHandle) = True Then Return False
            Return True
        End Get
    End Property

    '''<summary>Init with the DLL specified.</summary>
    Public Sub New(ByVal ippRoot As String)
        Me.New(System.IO.Path.Combine(ippRoot, "ipps.dll"), System.IO.Path.Combine(ippRoot, "ippvm.dll"), System.IO.Path.Combine(ippRoot, "ippi.dll"))
    End Sub

    '''<summary>Init with the DLL specified.</summary>
    Public Sub New(ByVal ippsDLL As String, ByVal ippvmDLL As String, ByVal ippiDLL As String)
        If System.IO.File.Exists(ippsDLL) And System.IO.File.Exists(ippvmDLL) And System.IO.File.Exists(ippiDLL) Then
            Try
                ChDir(System.IO.Path.GetDirectoryName(ippsDLL))
                ippsHandle = LoadLibrary(ippsDLL)
                ippvmHandle = LoadLibrary(ippvmDLL)
                ippiHandle = LoadLibrary(ippiDLL)
            Catch ex As Exception
                MyLoadError = ex.Message
                ippsHandle = Nothing
                ippvmHandle = Nothing
                ippiHandle = Nothing
            End Try
        End If
    End Sub

    Protected Overrides Sub Finalize()
        FreeLibrary(ippsHandle)
        MyBase.Finalize()
    End Sub

#End Region

#Region "Kernel32 Library Handling"

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function LoadLibrary(lpFileName As String) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function AddDllDirectory(Path As String) As Boolean
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True)> _
    Private Shared Function GetProcAddress(<System.Runtime.InteropServices.InAttribute> ByVal hModule As IntPtr, <System.Runtime.InteropServices.InAttribute, System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPStr)> ByVal lpProcName As String) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True, EntryPoint:="FreeLibrary")> _
    Private Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
    End Function

#End Region

    '''<summary>Class to handle GAC pinning and release.</summary>
    Private Class cPinHandler
        Private Pinned As New List(Of Runtime.InteropServices.GCHandle)
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin() As T) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(ArrayToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayToPin, 0)
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin(,) As T) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(ArrayToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayToPin, 0)
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin(,) As T, ByRef Offset As Integer) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(ArrayToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayToPin, Offset)
        End Function
        '''<summary>Release all pinned objects.</summary>
        Protected Overrides Sub Finalize()
            For Each PinnedObject As Runtime.InteropServices.GCHandle In Pinned
                PinnedObject.Free()
            Next PinnedObject
            MyBase.Finalize()
        End Sub
    End Class

    '''<summary>Region size</summary>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-structures-and-enumerators-1</remarks>
    Public Structure sIppiSize
        Public Width As Integer
        Public Height As Integer
        Public Sub New(ByVal W As Integer, ByVal H As Integer)
            Me.Width = W
            Me.Height = H
        End Sub
    End Structure

#Region "AdjustSize"
    Friend Shared Sub AdjustSize(Of InT, OutT)(ByRef Source() As InT, ByRef Target() As OutT)
        If Source.Length <> Target.Length Then ReDim Target(0 To Source.Length - 1)
    End Sub
    Friend Shared Sub AdjustSize(Of InT, OutT)(ByRef Source(,) As InT, ByRef Target(,) As OutT)
        Dim Adjust As Boolean = False
        If Source.GetUpperBound(0) <> Target.GetUpperBound(0) Then Adjust = True
        If Source.GetUpperBound(1) <> Target.GetUpperBound(1) Then Adjust = True
        If Adjust Then ReDim Target(0 To Source.GetUpperBound(0), 0 To Source.GetUpperBound(1))
    End Sub
    Friend Shared Sub AdjustSize(Of OutT)(ByRef SourceLength As Integer, ByRef Target() As OutT)
        If SourceLength <> Target.Length Then ReDim Target(0 To SourceLength - 1)
    End Sub
#End Region

#Region "Delegates"
    Private Delegate Function CallSignature_Single_IntPtr_Integer(ByVal val As Single, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_IntPtr_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_Integer_IntPtr_Integer(ByVal pSrc As IntPtr, ByVal len As Integer, ByVal pMin As IntPtr, ByVal scaleFactor As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_IntPtr_Integer_IppRoundMode_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer, ByVal rndMode As IppRoundMode, ByVal scaleFactor As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_Integer_IntPtr_IntPtr(ByVal pSrc As IntPtr, ByVal len As Integer, ByVal pMin As IntPtr, ByVal pMax As IntPtr) As IppStatus
    Private Delegate Function CallSignature_Double_IntPtr_Integer(ByVal val As Double, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_Integer(ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_UInt16_IntPtr_Integer(ByVal val As UInt16, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_UInt32_IntPtr_Integer(ByVal val As UInt32, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function CallSignature_IntPtr_Integer_IppiSize(ByVal pSrcDst As IntPtr, ByVal iDst As Integer, ByVal roiSize As sIppiSize) As IppStatus
    Private Delegate Function CallSignature_IntPtr_Integer_IntPtr_Integer_IppiSize(ByVal pSrc As IntPtr, ByVal iSrc As Integer, ByVal pDst As IntPtr, ByVal iDst As Integer, ByVal roiSize As sIppiSize) As IppStatus
    Private Delegate Function CallSignature_IntPtr_IntPtr_IntPtr_Integer(ByVal pSrc1 As IntPtr, ByVal PSrc2 As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer) As IppStatus
#End Region

#Region "Enums"

    Public Enum IppRoundMode
        ippRndZero = 0
        ippRndNear = 1
    End Enum

    Public Enum IppStatus
        NotSupportedModeErr = -9999
        CpuNotSupportedErr = -9998
        ConvergeErr = -205
        SizeMatchMatrixErr = -204
        CountMatrixErr = -203
        RoiShiftMatrixErr = -202
        ResizeNoOperationErr = -201
        SrcDataErr = -200
        MaxLenHuffCodeErr = -199
        CodeLenTableErr = -198
        FreqTableErr = -197
        IncompleteContextErr = -196
        SingularErr = -195
        SparseErr = -194
        BitOffsetErr = -193
        QPErr = -192
        VLCErr = -191
        RegExpOptionsErr = -190
        RegExpErr = -189
        RegExpMatchLimitErr = -188
        RegExpQuantifierErr = -187
        RegExpGroupingErr = -186
        RegExpBackRefErr = -185
        RegExpChClassErr = -184
        RegExpMetaChErr = -183
        StrideMatrixErr = -182
        CTRSizeErr = -181
        JPEG2KCodeBlockIsNotAttached = -180
        NotPosDefErr = -179
        EphemeralKeyErr = -178
        MessageErr = -177
        ShareKeyErr = -176
        IvalidPublicKey = -175
        IvalidPrivateKey = -174
        OutOfECErr = -173
        ECCInvalidFlagErr = -172
        MP3FrameHeaderErr = -171
        MP3SideInfoErr = -170
        BlockStepErr = -169
        MBStepErr = -168
        AacPrgNumErr = -167
        AacSectCbErr = -166
        AacSfValErr = -164
        AacCoefValErr = -163
        AacMaxSfbErr = -162
        AacPredSfbErr = -161
        AacPlsDataErr = -160
        AacGainCtrErr = -159
        AacSectErr = -158
        AacTnsNumFiltErr = -157
        AacTnsLenErr = -156
        AacTnsOrderErr = -155
        AacTnsCoefResErr = -154
        AacTnsCoefErr = -153
        AacTnsDirectErr = -152
        AacTnsProfileErr = -151
        AacErr = -150
        AacBitOffsetErr = -149
        AacAdtsSyncWordErr = -148
        AacSmplRateIdxErr = -147
        AacWinLenErr = -146
        AacWinGrpErr = -145
        AacWinSeqErr = -144
        AacComWinErr = -143
        AacStereoMaskErr = -142
        AacChanErr = -141
        AacMonoStereoErr = -140
        AacStereoLayerErr = -139
        AacMonoLayerErr = -138
        AacScalableErr = -137
        AacObjTypeErr = -136
        AacWinShapeErr = -135
        AacPcmModeErr = -134
        VLCUsrTblHeaderErr = -133
        VLCUsrTblUnsupportedFmtErr = -132
        VLCUsrTblEscAlgTypeErr = -131
        VLCUsrTblEscCodeLengthErr = -130
        VLCUsrTblCodeLengthErr = -129
        VLCInternalTblErr = -128
        VLCInputDataErr = -127
        VLCAACEscCodeLengthErr = -126
        NoiseRangeErr = -125
        UnderRunErr = -124
        PaddingErr = -123
        CFBSizeErr = -122
        PaddingSchemeErr = -121
        InvalidCryptoKeyErr = -120
        LengthErr = -119
        BadModulusErr = -118
        LPCCalcErr = -117
        RCCalcErr = -116
        IncorrectLSPErr = -115
        NoRootFoundErr = -114
        JPEG2KBadPassNumber = -113
        JPEG2KDamagedCodeBlock = -112
        H263CBPYCodeErr = -111
        H263MCBPCInterCodeErr = -110
        H263MCBPCIntraCodeErr = -109
        NotEvenStepErr = -108
        HistoNofLevelsErr = -107
        LUTNofLevelsErr = -106
        MP4BitOffsetErr = -105
        MP4QPErr = -104
        MP4BlockIdxErr = -103
        MP4BlockTypeErr = -102
        MP4MVCodeErr = -101
        MP4VLCCodeErr = -100
        MP4DCCodeErr = -99
        MP4FcodeErr = -98
        MP4AlignErr = -97
        MP4TempDiffErr = -96
        MP4BlockSizeErr = -95
        MP4ZeroBABErr = -94
        MP4PredDirErr = -93
        MP4BitsPerPixelErr = -92
        MP4VideoCompModeErr = -91
        MP4LinearModeErr = -90
        H263PredModeErr = -83
        H263BlockStepErr = -82
        H263MBStepErr = -81
        H263FrameWidthErr = -80
        H263FrameHeightErr = -79
        H263ExpandPelsErr = -78
        H263PlaneStepErr = -77
        H263QuantErr = -76
        H263MVCodeErr = -75
        H263VLCCodeErr = -74
        H263DCCodeErr = -73
        H263ZigzagLenErr = -72
        FBankFreqErr = -71
        FBankFlagErr = -70
        FBankErr = -69
        NegOccErr = -67
        CdbkFlagErr = -66
        SVDCnvgErr = -65
        JPEGHuffTableErr = -64
        JPEGDCTRangeErr = -63
        JPEGOutOfBufErr = -62
        DrawTextErr = -61
        ChannelOrderErr = -60
        ZeroMaskValuesErr = -59
        QuadErr = -58
        RectErr = -57
        CoeffErr = -56
        NoiseValErr = -55
        DitherLevelsErr = -54
        NumChannelsErr = -53
        COIErr = -52
        DivisorErr = -51
        AlphaTypeErr = -50
        GammaRangeErr = -49
        GrayCoefSumErr = -48
        ChannelErr = -47
        ToneMagnErr = -46
        ToneFreqErr = -45
        TonePhaseErr = -44
        TrnglMagnErr = -43
        TrnglFreqErr = -42
        TrnglPhaseErr = -41
        TrnglAsymErr = -40
        HugeWinErr = -39
        JaehneErr = -38
        StrideErr = -37
        EpsValErr = -36
        WtOffsetErr = -35
        AnchorErr = -34
        MaskSizeErr = -33
        ShiftErr = -32
        SampleFactorErr = -31
        SamplePhaseErr = -30
        FIRMRFactorErr = -29
        FIRMRPhaseErr = -28
        RelFreqErr = -27
        FIRLenErr = -26
        IIROrderErr = -25
        DlyLineIndexErr = -24
        ResizeFactorErr = -23
        InterpolationErr = -22
        MirrorFlipErr = -21
        Moment00ZeroErr = -20
        ThreshNegLevelErr = -19
        ThresholdErr = -18
        ContextMatchErr = -17
        FftFlagErr = -16
        FftOrderErr = -15
        StepErr = -14
        ScaleRangeErr = -13
        DataTypeErr = -12
        OutOfRangeErr = -11
        DivByZeroErr = -10
        MemAllocErr = -9
        NullPtrErr = -8
        RangeErr = -7
        SizeErr = -6
        BadArgErr = -5
        NoMemErr = -4
        SAReservedErr3 = -3
        Err = -2
        SAReservedErr1 = -1
        NoErr = 0
        NoOperation = 1
        MisalignedBuf = 2
        SqrtNegArg = 3
        InvZero = 4
        EvenMedianMaskSize = 5
        DivByZero = 6
        LnZeroArg = 7
        LnNegArg = 8
        NanArg = 9
        JPEGMarker = 10
        ResFloor = 11
        Overflow = 12
        LSFLow = 13
        LSFHigh = 14
        LSFLowAndHigh = 15
        ZeroOcc = 16
        Underflow = 17
        Singularity = 18
        Domain = 19
        NonIntelCpu = 20
        CpuMismatch = 21
        NoIppFunctionFound = 22
        DllNotFoundBestUsed = 23
        NoOperationInDll = 24
        InsufficientEntropy = 25
        OvermuchStrings = 26
        OverlongString = 27
        AffineQuadChanged = 28
        WrongIntersectROI = 29
        WrongIntersectQuad = 30
        SmallerCodebook = 31
        SrcSizeLessExpected = 32
        DstSizeLessExpected = 33
        StreamEnd = 34
        DoubleSize = 35
        NotSupportedCpu = 36
        UnknownCacheSize = 37
        SymKernelExpected = 38
    End Enum

#End Region

    '''<summary>Initializes a vector to zero.</summary>
    '''<param name="Src"></param>
    '''<returns>IPP status.</returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-zero</remarks>
    Public Function Zero(ByRef Src(,) As UInt32) As IppStatus
        Dim FunctionName As String = "ippsZero_32s"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Src), Src.Length), IppStatus)
    End Function

    'Add
    Public Function Add(ByRef Src(,) As UInt32, ByRef SrcDst(,) As UInt32) As IppStatus
        Dim FunctionName As String = "ippsAdd_32u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(Src, SrcDst)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(SrcDst), Src.Length), IppStatus)
    End Function

    'AddC
    Public Function AddC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim FunctionName As String = "ippsAddC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Single_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    'AddC
    Public Function AddC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsAddC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Double_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    'SubC
    Public Function SubC(ByRef Vector(,) As Single, ByVal SubVal As Single) As IppStatus
        Dim FunctionName As String = "ippsSubC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Single_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(SubVal, Pinner.Pin(Vector), Vector.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'MulC
    Public Function MulC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim FunctionName As String = "ippsMulC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Single_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsMulC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Single_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Int32, ByRef ScaleFactor As Int32) As IppStatus
        Dim FunctionName As String = "ippsMulC_32s_ISfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Double_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length, 0), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Mul
    Public Function Mul(ByRef ArraySrc(,) As Double, ByRef ArraySrcDst(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsMul_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(ArraySrc), Pinner.Pin(ArraySrcDst), ArraySrc.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Sum
    '''<summary>Computes the sum of the elements of a vector.</summary>
    '''<param name="ArraySrc">Array to calculate sum from.</param>
    '''<param name="TotalSum">Sum of all elements.</param>
    Public Function Sum(ByRef ArraySrc(,) As Short, ByRef TotalSum As Integer) As IppStatus
        Dim FunctionName As String = "ippsSum_16s32s_Sfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Dim ArrayDst(0) As Integer
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(ArraySrc), ArraySrc.Length, Pinner.Pin(ArrayDst), 0), IppStatus)
        TotalSum = ArrayDst(0)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    '''<summary>Divides each element of a vector by a constant value.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-divc</remarks>
    Public Function DivC(ByRef Array() As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsDivC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Double_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '''<summary>Divides each element of a vector by a constant value.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-divc</remarks>
    Public Function DivC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsDivC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_Double_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '''<summary>Divides each element of a vector by a constant value.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>Function does not exist in IPP.</remarks>
    Public Function DivC(ByRef Array(,) As UInt32, ByRef ScaleFactor As UInt32) As IppStatus
        For Idx1 As Integer = 0 To Array.GetUpperBound(0)
            For Idx2 As Integer = 0 To Array.GetUpperBound(1)
                Array(Idx1, Idx2) = CUInt(Array(Idx1, Idx2) / ScaleFactor)
            Next Idx2
        Next Idx1
        Return IppStatus.NoErr
    End Function

    '==========================================================================================================================================================

    ''' <summary>Convert UInt16 to UInt32 (using RealToCplx as the convert function does not exist).</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-realtocplx</remarks>
    Public Function Convert(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut(,) As UInt32) As IppStatus
        Dim FunctionName As String = "ippsRealToCplx_16s"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(ArrayIn, ArrayOut)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), IntPtr.Zero, Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Short, ByVal RoundMode As IppRoundMode, ByVal ScaleFactor As Integer) As IppStatus
        Dim FunctionName As String = "ippsConvert_64f16s_Sfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer_IppRoundMode_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(ArrayIn, ArrayOut)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length, RoundMode, ScaleFactor), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Single) As IppStatus
        Dim FunctionName As String = "ippsConvert_64f32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(ArrayIn, ArrayOut)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Single) As IppStatus
        Dim FunctionName As String = "ippsConvert_32s32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(Src, Dst)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsConvert_32s64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        AdjustSize(Src, Dst)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'MinMax
    Public Function MinMax(ByRef Array() As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim FunctionName As String = "ippsMinMax_32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_IntPtr))
        Dim Pinner As New cPinHandler
        Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MinMax
    Public Function MinMax(ByRef Array(,) As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim FunctionName As String = "ippsMinMax_32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_IntPtr))
        Dim Pinner As New cPinHandler
        Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MinMax
    Public Function MinMax(ByRef Array(,) As Double, ByRef Minimum As Double, ByRef Maximum As Double) As IppStatus
        Dim FunctionName As String = "ippsMinMax_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_IntPtr))
        Dim Pinner As New cPinHandler
        Dim TempVal1(0) As Double : Dim TempVal2(0) As Double
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MinMax
    Public Function MinMax(ByRef Array() As Double, ByRef Minimum As Double, ByRef Maximum As Double) As IppStatus
        Dim FunctionName As String = "ippsMinMax_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_IntPtr))
        Dim Pinner As New cPinHandler
        Dim TempVal1(0) As Double : Dim TempVal2(0) As Double
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'Max (derived)
    Public Function Max(ByRef Array() As Double) As Double
        Dim Minimum As Double = Double.NaN
        Dim Maximum As Double = Double.NaN
        If MinMax(Array, Minimum, Maximum) = IppStatus.NoErr Then Return Maximum Else Return Double.NaN
    End Function

    '==========================================================================================================================================================

    'MaxIdx
    Public Function MaxIndx(ByRef Array(,) As Double, ByRef Maximum As Double, ByRef MaximumIdx As Integer) As IppStatus
        Dim FunctionName As String = "ippsMaxIndx_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_IntPtr))
        Dim Pinner As New cPinHandler
        Dim TempVal1(0) As Double : Dim TempVal2(0) As Integer
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
        Maximum = TempVal1(0) : MaximumIdx = TempVal2(0)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    'Sqr
    Public Function Sqr(ByRef Array(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSqr_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Sqrt
    Public Function Sqrt(ByRef Array(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSqrt_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Copy
    Public Function Copy(ByRef Vector As Double(,)) As Double(,)
        Dim FunctionName As String = "ippsCopy_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Dim RetVal(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Double
        Caller.DynamicInvoke(Pinner.Pin(Vector), Pinner.Pin(RetVal), RetVal.Length)
        Return RetVal
    End Function

    'Copy
    Public Function Copy(ByRef Vector As Int32(,)) As Int32(,)
        Dim FunctionName As String = "ippsCopy_32s"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Dim RetVal(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Int32
        Caller.DynamicInvoke(Pinner.Pin(Vector), Pinner.Pin(RetVal), RetVal.Length)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    ''' <summary>SwapBytes (used to swap bytes read via ReadBytes and convert them "in-memory-direct")</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Src() As Byte, ByRef Dst(,) As UInt16) As IppStatus
        Dim FunctionName As String = "ippsSwapBytes_16u"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length \ 2), IppStatus)
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As UInt16) As IppStatus
        Dim FunctionName As String = "ippsSwapBytes_16u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As UInt32) As IppStatus
        Dim FunctionName As String = "ippsSwapBytes_32u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As Int32) As IppStatus
        Dim FunctionName As String = "ippsSwapBytes_32u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    ''' <summary>Transpose</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef ArrayIn() As Byte, ByRef ArrayOut(,) As UInt16) As IppStatus
        Dim FunctionName As String = "ippiTranspose_16u_C1R"
        Dim FunPtr As IntPtr = GetProcAddress(ippiHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_Integer_IppiSize))
        Dim Pinner As New cPinHandler
        Dim srcStep As Integer = 2 * (ArrayOut.GetUpperBound(0) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the source image.
        Dim dstStep As Integer = 2 * (ArrayOut.GetUpperBound(1) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the destination image.
        Dim ROI As New sIppiSize(srcStep \ 2, dstStep \ 2)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
    End Function

    ''' <summary>Transpose</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef Array(,) As UInt16) As IppStatus
        Dim FunctionName As String = "ippiTranspose_32u_C1IR"
        Dim FunPtr As IntPtr = GetProcAddress(ippiHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IppiSize))
        Dim Pinner As New cPinHandler
        Dim srcDstStep As Integer = 2 * (Array.GetUpperBound(1) + 1)
        Dim ROI As sIppiSize = GetFullROI(Array)
        Return CType(Caller.DynamicInvoke(Pinner.Pin(Array), srcDstStep, ROI), IppStatus)
    End Function

    '==========================================================================================================================================================

    ''' <summary>Computes the bitwise XOR of a scalar value and each element of a vector.</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-xorc</remarks>
    Public Function XorC(ByRef Array(,) As UInt16, ByVal Value As UInt16) As IppStatus
        Dim FunctionName As String = "ippsXorC_16u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_UInt16_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Value, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    ''' <summary>Computes the bitwise XOR of a scalar value and each element of a vector.</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-xorc</remarks>
    Public Function XorC(ByRef Array(,) As UInt32, ByVal Value As UInt32) As IppStatus
        Dim FunctionName As String = "ippsXorC_32u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_UInt32_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Return CType(Caller.DynamicInvoke(Value, Pinner.Pin(Array), Array.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Sin
    Public Function Sin(ByRef ArrayIn(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSin_64f_A53"
        Dim FunPtr As IntPtr = GetProcAddress(ippvmHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        Dim InPlace(,) As Double = {}
        AdjustSize(ArrayIn, InPlace)
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(InPlace), ArrayIn.Length), IppStatus)
        ArrayIn = Copy(InPlace)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    '''<summary>Sort descending.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-sortascend-sortdescend</remarks>
    Public Function SortDescend(ByRef Array() As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim FunctionName As String = "ippsSortAscend_16u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        Return RetVal
    End Function

    '''<summary>Sort descending.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-sortascend-sortdescend</remarks>
    Public Function SortDescend(ByRef Array(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim FunctionName As String = "ippsSortDescend_16u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    '''<summary>Interleaved copy - copy 1 bayer channel of RGGB to a 1/4 size image.</summary>
    '''<param name="ArrayIn"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-copy-1</remarks>
    Public Function CopyPixel(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut(,) As UInt16, ByVal Offset As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim FunctionName As String = "ippiCopy_16u_C4C1R"
        Dim FunPtr As IntPtr = GetProcAddress(ippiHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_Integer_IppiSize))
        Dim Pinner As New cPinHandler
        ReDim ArrayOut(((ArrayIn.GetUpperBound(0) + 1) \ 2) - 1, ((ArrayIn.GetUpperBound(1) + 1) \ 2) - 1)
        Dim ROI As sIppiSize = GetFullROI(ArrayIn)
        RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn, Offset), 8, Pinner.Pin(ArrayOut), 2, ROI), IppStatus)
        Return RetVal
    End Function

    '==========================================================================================================================================================

    Private Function GetFullROI(ByRef ArrayIn(,) As UInt16) As sIppiSize
        Return New sIppiSize(2 * (ArrayIn.GetUpperBound(0) + 1), 2 * (ArrayIn.GetUpperBound(1) + 1))
    End Function

    '==========================================================================================================================================================
    ' EXTENDED FUNCTIONS
    '==========================================================================================================================================================

    '''<summary>Computes the sum of the elements of a vector.</summary>
    '''<param name="ArraySrc">Array to calculate sum from.</param>
    '''<param name="TotalSum">Sum of all elements.</param>
    '''<remarks>Function ippsSum_16s32s_Sfs is calles with scaling factor 0 and sliced.</remarks>
    Public Function Sum(ByRef ArraySrc(,) As Short, ByRef TotalSum As Long) As IppStatus
        Dim FunctionName As String = "ippsSum_16s32s_Sfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(CallSignature_IntPtr_Integer_IntPtr_Integer))
        Dim Pinner As New cPinHandler
        TotalSum = 0
        Dim ArrayDst(0) As Integer
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim SliceSize As Integer = Short.MaxValue - 1

        Pinner.Pin(ArraySrc)
        Pinner.Pin(ArrayDst)
        Try
            For Idx As Integer = 0 To CInt(ArraySrc.LongLength) Step SliceSize
                Dim Length As Integer = SliceSize : If Idx + Length > CInt(ArraySrc.LongLength) Then Length = CInt(ArraySrc.LongLength) - Idx
                RetVal = CType(Caller.DynamicInvoke(System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArraySrc, Idx), Length, System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayDst, 0), 0), IppStatus)
                TotalSum += ArrayDst(0)
            Next Idx
        Catch ex As Exception
            TotalSum = 0
            Return IppStatus.Err
        End Try

        Return RetVal

    End Function

End Class

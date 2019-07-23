Option Explicit On
Option Strict On

'''<summary>Use this class to dynamically call any IPP DLL.</summary>
Partial Public Class cIntelIPP

#Region "Handles and constructors"

    '''<summary>Handle to the DLL.</summary>
    Private ippsHandle As IntPtr = Nothing
    Private ippvmHandle As IntPtr = Nothing

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
    Public Sub New(ByVal ippsDLL As String, ByVal ippvmDLL As String)
        If System.IO.File.Exists(ippsDLL) And System.IO.File.Exists(ippvmDLL) Then
            Try
                ChDir(System.IO.Path.GetDirectoryName(ippsDLL))
                ippsHandle = LoadLibrary(ippsDLL)
                ippvmHandle = LoadLibrary(ippvmDLL)
            Catch ex As Exception
                MyLoadError = ex.Message
                ippsHandle = Nothing
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

#Region "GetPtr"

    Friend Shared Function GetPtr(Of T)(ByRef Array() As T) As IntPtr
        Try
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(Array, 0)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Friend Shared Function GetPtr(Of T)(ByRef Array() As T, ByRef Offset As Integer) As IntPtr
        Try
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(Array, Offset)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Friend Shared Function GetPtr(Of T)(ByRef Array(,) As T) As IntPtr
        Try
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(Array, 0)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

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
    Private Delegate Function Call_Single_IntPtr_Integer(ByVal val As Single, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function Call_IntPtr_IntPtr_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function Call_IntPtr_IntPtr_Integer_IppRoundMode_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer, ByVal rndMode As IppRoundMode, ByVal scaleFactor As Integer) As IppStatus
    Private Delegate Function Call_IntPtr_Integer_IntPtr_IntPtr(ByVal pSrc As IntPtr, ByVal len As Integer, ByVal pMin As IntPtr, ByVal pMax As IntPtr) As IppStatus
    Private Delegate Function Call_Double_IntPtr_Integer(ByVal val As Double, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
    Private Delegate Function Call_IntPtr_Integer(ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
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

    'AddC - function does not exist ...
    Public Function Add(ByRef Src(,) As Int32, ByRef SrcDst(,) As Int32) As IppStatus
        AdjustSize(Src, SrcDst)
        For Idx1 As Integer = 0 To Src.GetUpperBound(0)
            For Idx2 As Integer = 0 To Src.GetUpperBound(1)
                SrcDst(Idx1, Idx2) = SrcDst(Idx1, Idx2) + Src(Idx1, Idx2)
            Next Idx2
        Next Idx1
        Return IppStatus.NoErr
    End Function

    'AddC
    Public Function AddC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim FunctionName As String = "ippsAddC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Single_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length), IppStatus)
    End Function

    'AddC
    Public Function AddC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsAddC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Double_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length), IppStatus)
    End Function

    'SubC
    Public Function SubC(ByRef Vector(,) As Single, ByVal SubVal As Single) As IppStatus
        Dim FunctionName As String = "ippsSubC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Single_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(SubVal, GetPtr(Vector), Vector.Length), IppStatus)
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim FunctionName As String = "ippsMulC_32f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Single_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length), IppStatus)
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsMulC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Single_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length), IppStatus)
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Int32, ByRef ScaleFactor As Int32) As IppStatus
        Dim FunctionName As String = "ippsMulC_32s_ISfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Double_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length, 0), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Mul
    Public Function Mul(ByRef ArraySrc(,) As Double, ByRef ArraySrcDst(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsMul_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(GetPtr(ArraySrc), GetPtr(ArraySrcDst), ArraySrc.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'DivC
    Public Function DivC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim FunctionName As String = "ippsDivC_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_Double_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(ScaleFactor, GetPtr(Array), Array.Length), IppStatus)
    End Function

    '==========================================================================================================================================================

    'Convert
    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Short, ByVal RoundMode As IppRoundMode, ByVal ScaleFactor As Integer) As IppStatus
        Dim FunctionName As String = "ippsConvert_64f16s_Sfs"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer_IppRoundMode_Integer))
        AdjustSize(ArrayIn, ArrayOut)
        Return CType(Caller.DynamicInvoke(GetPtr(ArrayIn), GetPtr(ArrayOut), ArrayIn.Length, RoundMode, ScaleFactor), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Single) As IppStatus
        Dim FunctionName As String = "ippsConvert_64f32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer))
        AdjustSize(ArrayIn, ArrayOut)
        Return CType(Caller.DynamicInvoke(GetPtr(ArrayIn), GetPtr(ArrayOut), ArrayIn.Length), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Single) As IppStatus
        Dim FunctionName As String = "ippsConvert_32s32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer))
        AdjustSize(Src, Dst)
        Return CType(Caller.DynamicInvoke(GetPtr(Src), GetPtr(Dst), Src.Length), IppStatus)
    End Function

    'Convert
    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsConvert_32s64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer))
        AdjustSize(Src, Dst)
        Return CType(Caller.DynamicInvoke(GetPtr(Src), GetPtr(Dst), Src.Length), IppStatus)
    End Function

    'MinMax
    Public Function MinMax(ByRef Array() As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim FunctionName As String = "ippsMinMax_32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer_IntPtr_IntPtr))
        Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length, GetPtr(TempVal1), GetPtr(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MinMax
    Public Function MinMax(ByRef Array(,) As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim FunctionName As String = "ippsMinMax_32f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer_IntPtr_IntPtr))
        Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length, GetPtr(TempVal1), GetPtr(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MinMax
    Public Function MinMax(ByRef Array(,) As Double, ByRef Minimum As Double, ByRef Maximum As Double) As IppStatus
        Dim FunctionName As String = "ippsMinMax_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer_IntPtr_IntPtr))
        Dim TempVal1(0) As Double : Dim TempVal2(0) As Double
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length, GetPtr(TempVal1), GetPtr(TempVal2)), IppStatus)
        Minimum = TempVal1(0) : Maximum = TempVal2(0)
        Return RetVal
    End Function

    'MaxIdx
    Public Function MaxIndx(ByRef Array(,) As Double, ByRef Maximum As Double, ByRef MaximumIdx As Integer) As IppStatus
        Dim FunctionName As String = "ippsMaxIndx_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer_IntPtr_IntPtr))
        Dim TempVal1(0) As Double : Dim TempVal2(0) As Integer
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length, GetPtr(TempVal1), GetPtr(TempVal2)), IppStatus)
        Maximum = TempVal1(0) : MaximumIdx = TempVal2(0)
        Return RetVal
    End Function

    'Sqr
    Public Function Sqr(ByRef Array(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSqr_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length), IppStatus)
    End Function

    'Sqrt
    Public Function Sqrt(ByRef Array(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSqrt_64f_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length), IppStatus)
    End Function

    'Copy
    Public Function Copy(ByRef Vector As Double(,)) As Double(,)
        Dim FunctionName As String = "ippsCopy_64f"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer))
        Dim RetVal(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Double
        Caller.DynamicInvoke(GetPtr(Vector), GetPtr(RetVal), RetVal.Length)
        Return RetVal
    End Function

    'Copy
    Public Function Copy(ByRef Vector As Int32(,)) As Int32(,)
        Dim FunctionName As String = "ippsCopy_32s"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer))
        Dim RetVal(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Int32
        Caller.DynamicInvoke(GetPtr(Vector), GetPtr(RetVal), RetVal.Length)
        Return RetVal
    End Function

    'SwapBytes 
    Public Function SwapBytes(ByRef Array(,) As Int32) As IppStatus
        Dim FunctionName As String = "ippsSwapBytes_32u_I"
        Dim FunPtr As IntPtr = GetProcAddress(ippsHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_Integer))
        Return CType(Caller.DynamicInvoke(GetPtr(Array), Array.Length), IppStatus)
    End Function



    Public Function Sin(ByRef ArrayIn(,) As Double) As IppStatus
        Dim FunctionName As String = "ippsSin_64f_A53"
        Dim FunPtr As IntPtr = GetProcAddress(ippvmHandle, FunctionName)
        Dim Caller As System.Delegate = Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(FunPtr, GetType(Call_IntPtr_IntPtr_Integer))
        Dim InPlace(,) As Double = {}
        AdjustSize(ArrayIn, InPlace)
        Dim RetVal As IppStatus = CType(Caller.DynamicInvoke(GetPtr(ArrayIn), GetPtr(InPlace), ArrayIn.Length), IppStatus)
        ArrayIn = Copy(InPlace)
        Return RetVal
    End Function

End Class

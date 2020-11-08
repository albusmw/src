Option Explicit On
Option Strict On

'''<summary>Use this class to dynamically call any IPP DLL.</summary>
'''<see cref=">https://software.intel.com/en-us/ipp-dev-guide"/>
'''<see cref=">https://software.intel.com/en-us/articles/descriptor-codes-and-parameters-for-ippi-functions"/>
Partial Public Class cIntelIPP

    Const UInt16Bytes As Integer = 2
    Const UInt32Bytes As Integer = 4

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

    '''<summary>Currently used IPP path.</summary>
    Public ReadOnly Property IPPPath() As String
        Get
            Return MyIPPPath
        End Get
    End Property
    Private MyIPPPath As String = String.Empty

    '''<summary>Init with the DLL specified.</summary>
    Public Sub New(ByVal ippRoot As String)
        Me.New(System.IO.Path.Combine(ippRoot, "ipps.dll"), System.IO.Path.Combine(ippRoot, "ippvm.dll"), System.IO.Path.Combine(ippRoot, "ippi.dll"))
    End Sub

    '''<summary>Init with the DLL specified.</summary>
    Public Sub New(ByVal ippsDLL As String, ByVal ippvmDLL As String, ByVal ippiDLL As String)
        If System.IO.File.Exists(ippsDLL) And System.IO.File.Exists(ippvmDLL) And System.IO.File.Exists(ippiDLL) Then
            Try
                Dim IPPPathUsed As String = System.IO.Path.GetDirectoryName(ippsDLL)
                ChDir(IPPPathUsed)
                ippsHandle = LoadLibrary(ippsDLL)
                ippvmHandle = LoadLibrary(ippvmDLL)
                ippiHandle = LoadLibrary(ippiDLL)
                MyIPPPath = IPPPathUsed
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
        FreeLibrary(ippvmHandle)
        FreeLibrary(ippiHandle)
        MyBase.Finalize()
    End Sub

    '''<summary>Call an Intel IPP Signal Processing function.</summary>
    Private Function CallIPPS(ByVal FunctionName As String, ByVal Signature As Type) As System.Delegate
        Return Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(GetProcAddress(ippsHandle, FunctionName), Signature)
    End Function

    '''<summary>Call an Intel IPP Image Processing function.</summary>
    Private Function CallIPPI(ByVal FunctionName As String, ByVal Signature As Type) As System.Delegate
        Return Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(GetProcAddress(ippiHandle, FunctionName), Signature)
    End Function

#End Region

#Region "Kernel32 Library Handling"

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function LoadLibrary(lpFileName As String) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function AddDllDirectory(Path As String) As Boolean
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True)>
    Private Shared Function GetProcAddress(<System.Runtime.InteropServices.InAttribute> ByVal hModule As IntPtr, <System.Runtime.InteropServices.InAttribute, System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPStr)> ByVal lpProcName As String) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True, EntryPoint:="FreeLibrary")>
    Private Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
    End Function

#End Region

#Region "cPinHandler"
    '''<summary>Class to handle GAC pinning and release.</summary>
    Public Class cPinHandler : Implements IDisposable
        Private Disposed As Boolean = False
        Private Pinned As New List(Of Runtime.InteropServices.GCHandle)
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(ByRef VariableToPin As Double) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(VariableToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return Runtime.InteropServices.GCHandle.ToIntPtr(Pinned(Pinned.Count - 1))
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin() As T) As IntPtr
            Return Pin(ArrayToPin, 0)
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin() As T, ByRef Offset As Integer) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(ArrayToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayToPin, Offset)
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin(,) As T) As IntPtr
            Return Pin(ArrayToPin, 0)
        End Function
        '''<summary>Pin the array and get the pointer.</summary>
        Public Function Pin(Of T)(ByRef ArrayToPin(,) As T, ByRef Offset As Integer) As IntPtr
            Pinned.Add(Runtime.InteropServices.GCHandle.Alloc(ArrayToPin, Runtime.InteropServices.GCHandleType.Pinned))
            Return System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(ArrayToPin, Offset)
        End Function
        Public Sub ForceUnpinAll()
            For Each PinnedObject As Runtime.InteropServices.GCHandle In Pinned
                PinnedObject.Free()
            Next PinnedObject
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Disposed Then
                If disposing Then
                    ForceUnpinAll()
                End If
            End If
            Disposed = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
    End Class
#End Region

#Region "Strcutures and enums"

    '''<summary>IPP data types.</summary>
    '''<remarks>Taken from ippbase.h.</remarks>
    Public Enum IppDataType
        ippUndef = -1
        ipp1u = 0
        ipp8u = 1
        ipp8uc = 2
        ipp8s = 3
        ipp8sc = 4
        ipp16u = 5
        ipp16uc = 6
        ipp16s = 7
        ipp16sc = 8
        ipp32u = 9
        ipp32uc = 10
        ipp32s = 11
        ipp32sc = 12
        ipp32f = 13
        ipp32fc = 14
        ipp64u = 15
        ipp64uc = 16
        ipp64s = 17
        ipp64sc = 18
        ipp64f = 19
        ipp64fc = 20
    End Enum

    '''<summary>IPP interpolation types.</summary>
    '''<remarks>Taken from ipptypes.h.</remarks>
    Public Enum IppiInterpolationType
        ippNearest = 1
        ippLinear = 2
        ippCubic = 6
        ippLanczos = 16
        ippHahn = 0
        ippSuper = 8
    End Enum

    '''<remarks>Taken from ipptypes.h.</remarks>
    Public Enum IppiWarpDirection
        ippWarpForward
        ippWarpBackward
    End Enum

    '''<remarks>Taken from ipptypes.h.</remarks>
    Public Enum IppiBorderType
        '''<remarks>Border is replicated from the edge pixels.</remarks>
        ippBorderRepl = 1
        ippBorderWrap = 2
        ippBorderMirror = 3 '* left border: 012... -> 21012... */
        ippBorderMirrorR = 4 '* left border: 012... -> 210012... */
        ippBorderDefault = 5
        '''<remarks>Values of all border pixels are set to constant.</remarks>
        ippBorderConst = 6
        '''<remarks>Outer pixels are not processed.</remarks>
        ippBorderTransp = 7
        'Flags to use source image memory pixels from outside of the border in particular directions */
        ippBorderInMemTop = &H10
        ippBorderInMemBottom = &H20
        ippBorderInMemLeft = &H40
        ippBorderInMemRight = &H80
        '''<remarks>Border is obtained from the source image pixels in memory.</remarks>
        ippBorderInMem = ippBorderInMemLeft Or ippBorderInMemTop Or ippBorderInMemRight Or ippBorderInMemBottom
        'Flags to use source image memory pixels from outside of the border for first stage only in multi-stage filters */
        ippBorderFirstStageInMemTop = &H100
        ippBorderFirstStageInMemBottom = &H200
        ippBorderFirstStageInMemLeft = &H400
        ippBorderFirstStageInMemRight = &H800
        ippBorderFirstStageInMem = ippBorderFirstStageInMemLeft Or ippBorderFirstStageInMemTop Or ippBorderFirstStageInMemRight Or ippBorderFirstStageInMemBottom
    End Enum

    '''<summary>Region size</summary>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-structures-and-enumerators-1</remarks>
    Public Structure IppiSize
        Public Width As Integer
        Public Height As Integer
        Public Sub New(ByVal W As Integer, ByVal H As Integer)
            Me.Width = W
            Me.Height = H
        End Sub
        Public Function [Get](ByVal W As Integer, ByVal H As Integer) As IppiSize
            Me.Width = W
            Me.Height = H
            Return Me
        End Function
    End Structure
#End Region

#Region "AdjustSize"
    Friend Shared Sub AdjustSize(Of InT, OutT)(ByRef Source() As InT, ByRef Target() As OutT)
        If Source.Length <> Target.Length Then ReDim Target(0 To Source.Length - 1)
    End Sub
    Friend Shared Sub AdjustSize(Of InT, OutT)(ByRef Source(,) As InT, ByRef Target(,) As OutT)
        Dim Adjust As Boolean = False
        If IsNothing(Target) = True Then Target = New OutT(,) {}
        If Source.GetUpperBound(0) <> Target.GetUpperBound(0) Then Adjust = True
        If Source.GetUpperBound(1) <> Target.GetUpperBound(1) Then Adjust = True
        If Adjust Then ReDim Target(0 To Source.GetUpperBound(0), 0 To Source.GetUpperBound(1))
    End Sub
    Friend Shared Sub AdjustSize(Of OutT)(ByRef SourceLength As Integer, ByRef Target() As OutT)
        If SourceLength <> Target.Length Then ReDim Target(0 To SourceLength - 1)
    End Sub
#End Region

#Region "Delegates"
    Private Class CallSignature
        Public Delegate Function Single_IntPtr_Integer(ByVal val As Single, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function IntPtr_IntPtr_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function IntPtr_Integer_IntPtr_Integer(ByVal pSrc As IntPtr, ByVal len As Integer, ByVal pMin As IntPtr, ByVal scaleFactor As Integer) As IppStatus
        Public Delegate Function IntPtr_IntPtr_Integer_IppRoundMode_Integer(ByVal pSrc As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer, ByVal rndMode As IppRoundMode, ByVal scaleFactor As Integer) As IppStatus
        Public Delegate Function IntPtr_Integer_IntPtr_IntPtr(ByVal pSrc As IntPtr, ByVal len As Integer, ByVal pMin As IntPtr, ByVal pMax As IntPtr) As IppStatus
        Public Delegate Function Double_IntPtr_Integer(ByVal val As Double, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function IntPtr_Integer(ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function UInt16_IntPtr_Integer(ByVal val As UInt16, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function UInt16_IntPtr_Integer_Integer(ByVal val As UInt16, ByVal pSrcDst As IntPtr, ByVal len As Integer, ByVal scaleFactor As Integer) As IppStatus
        Public Delegate Function UInt32_IntPtr_Integer(ByVal val As UInt32, ByVal pSrcDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function IntPtr_Integer_IppiSize(ByVal pSrcDst As IntPtr, ByVal iDst As Integer, ByVal roiSize As IppiSize) As IppStatus
        Public Delegate Function IntPtr_Integer_IntPtr_Integer_IppiSize(ByVal pSrc As IntPtr, ByVal iSrc As Integer, ByVal pDst As IntPtr, ByVal iDst As Integer, ByVal roiSize As IppiSize) As IppStatus
        Public Delegate Function IntPtr_IntPtr_IntPtr_Integer(ByVal pSrc1 As IntPtr, ByVal PSrc2 As IntPtr, ByVal pDst As IntPtr, ByVal len As Integer) As IppStatus
        Public Delegate Function Double_Double_Double_IntPtr(ByVal angle As Double, ByVal xShift As Double, ByVal yShift As Double, ByVal coeffs As IntPtr) As IppStatus
        Public Delegate Function sIppiSize_sIppiSize_eIppDataType_IntPtr_eIppiInterpolationType_eIppiWarpDirection_eIppiBorderType_IntPtr_IntPtr(ByVal srcSize As IppiSize, ByVal dstSize As IppiSize, ByVal dataType As IppDataType, ByVal coeffs As IntPtr, ByVal interpolation As IppiInterpolationType, ByVal direction As IppiWarpDirection, ByVal borderType As IppiBorderType, ByVal SpecSize As IntPtr, ByVal InitBufSize As IntPtr) As IppStatus
        Public Delegate Function sIppiSize_sIppiSize_eIppDataType_IntPtr_eIppiWarpDirection_Integer_eIppiBorderType_Double_Integer_IppiWrapSpec(ByVal srcSize As IppiSize, ByVal dstSize As IppiSize, ByVal dataType As IppDataType, ByVal coeffs As IntPtr, ByVal direction As IppiWarpDirection, ByVal numChannels As Integer, ByVal borderType As IppiBorderType, ByVal boarderValue As IntPtr, ByVal smoothEdge As Integer, ByVal Spec As IntPtr) As IppStatus
    End Class


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

    '________________________________________________________________________________
    'IPP search
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Shared Function PossiblePaths(ByVal MyPath As String) As List(Of String)
        Dim RetVal As New List(Of String)
        RetVal.Add(System.IO.Path.Combine(MyPath, "ipp"))
        RetVal.Add(System.IO.Path.Combine(MyPath, "..\ipp"))
        RetVal.Add("C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2020.0.166\windows\redist\intel64_win\ipp")
        RetVal.Add("C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2019.5.281\windows\redist\intel64\ipp")
        RetVal.Add("C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2019.1.144\windows\redist\intel64_win\ipp\")
        Return RetVal
    End Function

    '''<summary>Try to get an instance to the latest available Intel IPP DLL.</summary>
    Public Shared Function SearchDLLToUse(ByVal Paths As String(), ByVal LoadError As String) As String
        Dim TestInstance As cIntelIPP = Nothing
        Try
            For Each IPPRoot As String In Paths
                If System.IO.Directory.Exists(IPPRoot) = True Then
                    Try
                        TestInstance = New cIntelIPP(IPPRoot)
                        If TestInstance.DLLHandleValid = True Then
                            LoadError = String.Empty
                            Return IPPRoot
                        End If
                    Catch ex As Exception
                        MsgBox("IPP <" & IPPRoot & "> not found!")
                    End Try
                End If
            Next IPPRoot
            If IsNothing(TestInstance) = True Then
                LoadError = "IPP not initiated!"
            Else
                If TestInstance.DLLHandleValid = False Then
                    LoadError = "IPP not found!"
                End If
            End If
        Catch ex As Exception
            LoadError = "Generic error on loading IPP: <" & ex.Message & ">"
        End Try
        Return String.Empty
    End Function

    '________________________________________________________________________________
    'Zero
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    '''<summary>Initializes a vector to zero.</summary>
    '''<param name="Src"></param>
    '''<returns>IPP status.</returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-zero</remarks>
    Public Function Zero(ByRef Src(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsZero_32s", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Src.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Add
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Add(ByRef Src(,) As UInt32, ByRef SrcDst(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsAdd_32u_I", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(Src, SrcDst)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(SrcDst), Src.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'AddC
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function AddC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsAddC_32f_I", GetType(CallSignature.Single_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function AddC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsAddC_64f_I", GetType(CallSignature.Double_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'SubC
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '''<see cref="https://software.intel.com/en-us/ipp-dev-reference-zero"/>
    Public Function SubC(ByRef Vector(,) As UInt16, ByVal SubVal As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSubC_16u_ISfs", GetType(CallSignature.UInt16_IntPtr_Integer_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(SubVal, Pinner.Pin(Vector), Vector.Length, 0), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function SubC(ByRef Vector(,) As Single, ByVal SubVal As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSubC_32f_I", GetType(CallSignature.Single_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(SubVal, Pinner.Pin(Vector), Vector.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'MulC
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function MulC(ByRef Array(,) As Single, ByRef ScaleFactor As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMulC_32f_I", GetType(CallSignature.Single_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMulC_64f_I", GetType(CallSignature.Single_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    'MulC
    Public Function MulC(ByRef Array(,) As Int32, ByRef ScaleFactor As Int32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMulC_32s_ISfs", GetType(CallSignature.Double_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length, 0), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Mul
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Mul(ByRef ArraySrc(,) As Double, ByRef ArraySrcDst(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMul_64f_I", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArraySrc), Pinner.Pin(ArraySrcDst), ArraySrc.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Sum
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    '''<summary>Computes the sum of the elements of a vector.</summary>
    '''<param name="ArraySrc">Array to calculate sum from.</param>
    '''<param name="TotalSum">Sum of all elements.</param>
    Public Function Sum(ByRef ArraySrc(,) As Short, ByRef TotalSum As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSum_16s32s_Sfs", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer))
        Using Pinner As New cPinHandler
            Dim ArrayDst(0) As Integer
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArraySrc), ArraySrc.Length, Pinner.Pin(ArrayDst), 0), IppStatus)
            TotalSum = ArrayDst(0)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'DivC
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    '''<summary>Divides each element of a vector by a constant value.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-divc</remarks>
    Public Function DivC(ByRef Array() As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsDivC_64f_I", GetType(CallSignature.Double_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '''<summary>Divides each element of a vector by a constant value.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-divc</remarks>
    Public Function DivC(ByRef Array(,) As Double, ByRef ScaleFactor As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsDivC_64f_I", GetType(CallSignature.Double_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(ScaleFactor, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
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

    '________________________________________________________________________________
    'Convert (not a real IPP function but used to get from UInt16 to UInt32)
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    ''' <summary>Convert UInt16 to UInt32 (using RealToCplx as the convert function does not exist).</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-realtocplx</remarks>
    Public Function Convert(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsRealToCplx_16s", GetType(CallSignature.IntPtr_IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(ArrayIn, ArrayOut)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), IntPtr.Zero, Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Short, ByVal RoundMode As IppRoundMode, ByVal ScaleFactor As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsConvert_64f16s_Sfs", GetType(CallSignature.IntPtr_IntPtr_Integer_IppRoundMode_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(ArrayIn, ArrayOut)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length, RoundMode, ScaleFactor), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function Convert(ByRef ArrayIn(,) As Double, ByRef ArrayOut(,) As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsConvert_64f32f", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(ArrayIn, ArrayOut)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsConvert_32s32f", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(Src, Dst)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function Convert(ByRef Src(,) As Int32, ByRef Dst(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsConvert_32s64f", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(Src, Dst)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'MinMax
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function MinMax(ByRef Array(,) As UInt32, ByRef Minimum As UInt32, ByRef Maximum As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMinMax_32u", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim TempVal1(0) As UInt32 : Dim TempVal2(0) As UInt32
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
            Minimum = TempVal1(0) : Maximum = TempVal2(0)
        End Using : Return RetVal
    End Function

    Public Function MinMax(ByRef Array() As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMinMax_32f", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
            Minimum = TempVal1(0) : Maximum = TempVal2(0)
        End Using : Return RetVal
    End Function

    Public Function MinMax(ByRef Array(,) As Single, ByRef Minimum As Single, ByRef Maximum As Single) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMinMax_32f", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim TempVal1(0) As Single : Dim TempVal2(0) As Single
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(TempVal1), Pinner.Pin(TempVal2)), IppStatus)
            Minimum = TempVal1(0) : Maximum = TempVal2(0)
        End Using : Return RetVal
    End Function

    Public Function MinMax(ByRef Array(,) As Double, ByRef Minimum As Double, ByRef Maximum As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMinMax_64f", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim RefTypes(1) As Double
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(RefTypes, 0), Pinner.Pin(RefTypes, 1)), IppStatus)
            Minimum = RefTypes(0) : Maximum = RefTypes(1)
        End Using : Return RetVal
    End Function

    '''<summary>Returns the maximum and minimum values of a vector.</summary>
    '''<param name="Array">Source vector.</param>
    '''<param name="Minimum">Minimum value.</param>
    '''<param name="Maximum">Maximum value.</param>
    '''<see cref="https://software.intel.com/en-us/ipp-dev-reference-minmax"/>
    Public Function MinMax(ByRef Array() As Double, ByRef Minimum As Double, ByRef Maximum As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMinMax_64f", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim RefTypes(1) As Double
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(RefTypes, 0), Pinner.Pin(RefTypes, 1)), IppStatus)
            Minimum = RefTypes(0) : Maximum = RefTypes(1)
        End Using : Return RetVal
    End Function

    Public Function Min(ByRef Array() As Double) As Double
        Dim Minimum As Double = Double.NaN
        Dim Maximum As Double = Double.NaN
        If MinMax(Array, Minimum, Maximum) = IppStatus.NoErr Then Return Minimum Else Return Double.NaN
    End Function

    Public Function Max(ByRef Array() As Double) As Double
        Dim Minimum As Double = Double.NaN
        Dim Maximum As Double = Double.NaN
        If MinMax(Array, Minimum, Maximum) = IppStatus.NoErr Then Return Maximum Else Return Double.NaN
    End Function

    '________________________________________________________________________________
    'MaxIdx
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    '''<summary>Returns the maximum value of a vector and the index of the maximum element.</summary>
    '''<see cref="https://software.intel.com/en-us/ipp-dev-reference-maxindx"/>
    Public Function MaxIndx(ByRef Array(,) As Double, ByRef Maximum As Double, ByRef MaximumIdx As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsMaxIndx_64f", GetType(CallSignature.IntPtr_Integer_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim RefDbl(0) As Double : Dim RefInt(0) As Integer
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length, Pinner.Pin(RefDbl), Pinner.Pin(RefInt)), IppStatus)
            Maximum = RefDbl(0) : MaximumIdx = RefInt(0)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Sqr
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Sqr(ByRef Array(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSqr_64f_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Sqrt
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Sqrt(ByRef Array(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSqrt_64f_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Copy
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Copy(ByRef Vector As Double(,)) As Double(,)
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsCopy_64f", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Dim RetVector(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Double
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Vector), Pinner.Pin(RetVector), RetVector.Length), IppStatus)
            If RetVal <> IppStatus.NoErr Then Return New Double(,) {}
        End Using : Return RetVector
    End Function

    Public Function Copy(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut() As Byte) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsCopy_16s", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
        End Using : Return RetVal
    End Function

    Public Function Copy(ByRef ArrayIn(,) As UInt32, ByRef ArrayOut(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsCopy_32s", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            AdjustSize(ArrayIn, ArrayOut)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(ArrayOut), ArrayIn.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <see cref="https://software.intel.com/en-us/ipp-dev-reference-copy"/>
    Public Function Copy(ByRef Vector As Int32(,)) As Int32(,)
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsCopy_32s", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Dim RetVector(Vector.GetUpperBound(0), Vector.GetUpperBound(1)) As Int32
        Using Pinner As New cPinHandler
            Caller.DynamicInvoke(Pinner.Pin(Vector), Pinner.Pin(RetVector), RetVector.Length)
        End Using : Return RetVector
    End Function

    ''' <summary>Copy.</summary>
    ''' <param name="FirstIndexStart">First (left) matrix index start position (0-based).</param>
    ''' <param name="SecondIndexStart">Second (right) matrix index start position (0-based).</param>
    ''' <param name="SecondIndexRange">First (left) matrix index range to copy.</param>
    ''' <param name="FirstIndexRange">Second (right) matrix index range to copy.</param>
    ''' <see cref="https://software.intel.com/en-us/ipp-dev-reference-copy-1"/>
    Public Function Copy(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut(,) As UInt16, ByVal FirstIndexStart As Integer, ByVal SecondIndexStart As Integer, ByVal FirstIndexRange As Integer, ByVal SecondIndexRange As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim BytePerVal As Integer = 2
        Dim Caller As System.Delegate = CallIPPI("ippiCopy_16u_C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim ArrayInWidth As Integer = ArrayIn.GetUpperBound(1) + 1
            Dim ArrayInHeight As Integer = ArrayIn.GetUpperBound(0) + 1
            ReDim ArrayOut(CInt(FirstIndexRange - 1), CInt(SecondIndexRange - 1))
            Dim srcStep As Integer = BytePerVal * ArrayInWidth                                          'Distance, in bytes, between the starting points of consecutive lines in the source image.
            Dim dstStep As Integer = BytePerVal * SecondIndexRange                                       'Distance, in bytes, between the starting points of consecutive lines in the destination image.
            Dim FirstValue As Integer = CInt(SecondIndexStart + (FirstIndexStart * ArrayInWidth))
            Dim ROI As New IppiSize(SecondIndexRange, FirstIndexRange)                                 'ROI [element index span - not depending on data format!]
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn, FirstValue), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    '''<summary>Interleaved copy - copy 1 bayer channel of RGGB to a 1/4 size image.</summary>
    '''<param name="ArrayIn"></param>
    '''<returns>IppStatus error code.</returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-copy-1</remarks>
    Public Function CopyPixel(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut(,) As UInt16, ByVal Offset As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiCopy_16u_C4C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            ReDim ArrayOut(((ArrayIn.GetUpperBound(0) + 1) \ 2) - 1, ((ArrayIn.GetUpperBound(1) + 1) \ 2) - 1)
            Dim ROI As IppiSize = GetFullROI(ArrayIn)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn, Offset), 8, Pinner.Pin(ArrayOut), 2, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'SwapBytes
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    ''' <summary>SwapBytes (used to swap bytes read via ReadBytes and convert them "in-memory-direct")</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Src() As Byte, ByRef Dst(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_16u", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length \ 2), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes (used to swap bytes read via ReadBytes and convert them "in-memory-direct")</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Src() As UInt16, ByRef Dst() As Byte) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_16u", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length \ 2), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes (used to swap bytes read via ReadBytes and convert them "in-memory-direct")</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Src(,) As UInt16, ByRef Dst() As Byte) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_16u", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Src), Pinner.Pin(Dst), Src.Length \ 2), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array() As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_16u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_16u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_32u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>SwapBytes</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-swapbytes</remarks>
    Public Function SwapBytes(ByRef Array(,) As Int32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSwapBytes_32u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Transpose
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    ''' <summary>Transpose - the destination image is obtained from the source image by transforming the columns to the rows.</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef ArrayIn() As Byte, ByRef ArrayOut(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiTranspose_16u_C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim srcStep As Integer = UInt16Bytes * (ArrayOut.GetUpperBound(0) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the source image.
            Dim dstStep As Integer = UInt16Bytes * (ArrayOut.GetUpperBound(1) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the destination image.
            Dim ROI As New IppiSize(srcStep \ UInt16Bytes, dstStep \ UInt16Bytes)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Transpose - the destination image is obtained from the source image by transforming the columns to the rows.</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef ArrayIn(,) As UInt16, ByRef ArrayOut() As Byte) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiTranspose_16u_C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim srcStep As Integer = UInt16Bytes * (ArrayIn.GetUpperBound(1) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the source image.
            Dim dstStep As Integer = UInt16Bytes * (ArrayIn.GetUpperBound(0) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the destination image.
            Dim ROI As New IppiSize(srcStep \ UInt16Bytes, dstStep \ UInt16Bytes)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Transpose - the destination image is obtained from the source image by transforming the columns to the rows.</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef ArrayIn() As Byte, ByRef ArrayOut(,) As Int32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiTranspose_32s_C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim srcStep As Integer = UInt32Bytes * (ArrayOut.GetUpperBound(0) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the source image.
            Dim dstStep As Integer = UInt32Bytes * (ArrayOut.GetUpperBound(1) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the destination image.
            Dim ROI As New IppiSize(srcStep \ UInt32Bytes, dstStep \ UInt32Bytes)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Transpose - the destination image is obtained from the source image by transforming the columns to the rows.</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef ArrayIn() As Byte, ByRef ArrayOut(,) As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiTranspose_32s_C1R", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim srcStep As Integer = UInt32Bytes * (ArrayOut.GetUpperBound(0) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the source image.
            Dim dstStep As Integer = UInt32Bytes * (ArrayOut.GetUpperBound(1) + 1)     'Distance, in bytes, between the starting points of consecutive lines in the destination image.
            Dim ROI As New IppiSize(srcStep \ UInt32Bytes, dstStep \ UInt32Bytes)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), srcStep, Pinner.Pin(ArrayOut), dstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Transpose - the destination image is obtained from the source image by transforming the columns to the rows.</summary>
    ''' <returns>IppStatus error code.</returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-transpose</remarks>
    Public Function Transpose(ByRef Array(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiTranspose_16u_C1IR", GetType(CallSignature.IntPtr_Integer_IppiSize))
        Using Pinner As New cPinHandler
            Dim srcDstStep As Integer = UInt16Bytes * (Array.GetUpperBound(1) + 1)
            Dim ROI As IppiSize = GetFullROI(Array)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), srcDstStep, ROI), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'XorC
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    ''' <summary>Computes the bitwise XOR of a scalar value and each element of a vector.</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-xorc</remarks>
    Public Function XorC(ByRef Array() As UInt16, ByVal Value As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsXorC_16u_I", GetType(CallSignature.UInt16_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Value, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Computes the bitwise XOR of a scalar value and each element of a vector.</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-xorc</remarks>
    Public Function XorC(ByRef Array(,) As UInt16, ByVal Value As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsXorC_16u_I", GetType(CallSignature.UInt16_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Value, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    ''' <summary>Computes the bitwise XOR of a scalar value and each element of a vector.</summary>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-xorc</remarks>
    Public Function XorC(ByRef Array(,) As UInt32, ByVal Value As UInt32) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsXorC_32u_I", GetType(CallSignature.UInt32_IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Value, Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'Sin
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    Public Function Sin(ByRef ArrayIn(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSin_64f_A53", GetType(CallSignature.IntPtr_IntPtr_Integer))
        Using Pinner As New cPinHandler
            Dim InPlace(,) As Double = {}
            AdjustSize(ArrayIn, InPlace)
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(ArrayIn), Pinner.Pin(InPlace), ArrayIn.Length), IppStatus)
            ArrayIn = Copy(InPlace)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'SortDescend
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    '''<summary>Sort descending.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-sortascend-sortdescend</remarks>
    Public Function SortDescend(ByRef Array() As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSortAscend_16u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '''<summary>Sort descending.</summary>
    '''<param name="Array"></param>
    '''<returns></returns>
    '''<remarks>https://software.intel.com/en-us/ipp-dev-reference-sortascend-sortdescend</remarks>
    Public Function SortDescend(ByRef Array(,) As UInt16) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSortDescend_16u_I", GetType(CallSignature.IntPtr_Integer))
        Using Pinner As New cPinHandler
            RetVal = CType(Caller.DynamicInvoke(Pinner.Pin(Array), Array.Length), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'GetRotateTransform
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    ''' <summary>Computes the affine coefficients for the rotation transform.</summary>
    ''' <param name="Angle">Angle of rotation, in degrees. The source image is rotated counterclockwise around the origin (0, 0).</param>
    ''' <param name="xShift">Shift along horizontal (x) axis that is performed after rotation.</param>
    ''' <param name="yShift">Shift along vertical (y) axis that is performed after rotation.</param>
    ''' <param name="coeffs">Computed affine transform coefficients for the given rotation parameters.</param>
    ''' <returns></returns>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-getrotatetransform</remarks>
    Public Function GetRotateTransform(ByVal Angle As Double, ByVal xShift As Double, ByVal yShift As Double, ByRef coeffs(,) As Double) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiGetRotateTransform", GetType(CallSignature.Double_Double_Double_IntPtr))
        Using Pinner As New cPinHandler
            ReDim coeffs(1, 2)
            RetVal = CType(Caller.DynamicInvoke(Angle, xShift, yShift, Pinner.Pin(coeffs)), IppStatus)
        End Using : Return RetVal
    End Function

    '________________________________________________________________________________
    'WarpAffineGetSize
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

    ''' <summary>Computes the size of the specification structure and the size of the external work buffer for the warp affine transform.</summary>
    ''' <param name="srcSize">Size of the source image, in pixels.</param>
    ''' <param name="dstSize">Size of the destination image, in pixels.</param>
    ''' <param name="dataType">Data type of the source and destination images. Supported values: ipp8u, ipp16u, ipp16s, ipp32f, and ipp64f.</param>
    ''' <param name="coeffs">Coefficients for the affine transform.</param>
    ''' <param name="interpolation">Interpolation method. Supported values: ippNearest, ippLinear and ippCubic.</param>
    ''' <param name="direction">Transformation direction.</param>
    ''' <param name="borderType">Type of border - supported: ippBorderConst, ippBorderRepl, ippBorderTransp, ippBorderInMem. Mixed borders are also supported.</param>
    ''' <param name="pSpecSize">Pointer to the size, in bytes, of the specification structure.</param>
    ''' <param name="pInitBufSize">Pointer to the size, in bytes, of the temporary buffer.</param>
    ''' <remarks>https://software.intel.com/en-us/ipp-dev-reference-warpaffinegetsize</remarks>
    Public Function WarpAffineGetSize(ByVal srcSize As IppiSize, ByVal dstSize As IppiSize, ByVal dataType As IppDataType, ByVal coeffs(,) As Double, ByVal interpolation As IppiInterpolationType, ByVal direction As IppiWarpDirection, ByVal borderType As IppiBorderType, ByRef pSpecSize As Integer, ByRef pInitBufSize As Integer) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiWarpAffineGetSize", GetType(CallSignature.sIppiSize_sIppiSize_eIppDataType_IntPtr_eIppiInterpolationType_eIppiWarpDirection_eIppiBorderType_IntPtr_IntPtr))
        Using Pinner As New cPinHandler
            Dim RefTypes(1) As Integer
            RetVal = CType(Caller.DynamicInvoke(srcSize, dstSize, dataType, Pinner.Pin(coeffs), interpolation, direction, borderType, Pinner.Pin(RefTypes, 0), Pinner.Pin(RefTypes, 1)), IppStatus)
            pSpecSize = RefTypes(0)
            pInitBufSize = RefTypes(1)
        End Using : Return RetVal
    End Function

    ''' <summary>   
    ''' 
    ''' </summary>
    ''' <param name="srcSize"></param>
    ''' <param name="dstSize"></param>
    ''' <param name="dataType"></param>
    ''' <param name="coeffs"></param>
    ''' <param name="direction"></param>
    ''' <param name="numChannels"></param>
    ''' <param name="borderType"></param>
    ''' <param name="borderValue"></param>
    ''' <param name="smoothEdge">Flag for edge smoothing. Supported values: 0: transformation without edge smoothing / 1: transformation with edge smoothing. This feature Is supported only for the ippBorderTransp And ippBorderInMem border types.</param>
    ''' <param name="pSpec"></param>
    ''' <returns></returns>
    Public Function WarpAffineLinearInit(ByVal srcSize As IppiSize, ByVal dstSize As IppiSize, ByVal dataType As IppDataType, ByVal coeffs(,) As Double, ByVal direction As IppiWarpDirection, ByVal numChannels As Integer, ByVal borderType As IppiBorderType, ByVal borderValue As Double, ByVal smoothEdge As Integer, ByRef pSpec() As Byte) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPI("ippiWarpAffineLinearInit", GetType(CallSignature.sIppiSize_sIppiSize_eIppDataType_IntPtr_eIppiWarpDirection_Integer_eIppiBorderType_Double_Integer_IppiWrapSpec))
        Using Pinner As New cPinHandler
            Dim RefTypes_dbl(0) As Double
            RetVal = CType(Caller.DynamicInvoke(srcSize, dstSize, dataType, Pinner.Pin(coeffs), direction, numChannels, borderType, Pinner.Pin(RefTypes_dbl, 0), 0, Pinner.Pin(pSpec, 0)), IppStatus)
            borderValue = RefTypes_dbl(0)
        End Using : Return RetVal
    End Function

    '================================================================================
    'EXTENDED FUNCTIONS

    Private Function GetFullROI(ByRef ArrayIn(,) As UInt16) As IppiSize
        Dim BytesPerVal As Integer = 2
        Return New IppiSize(BytesPerVal * (ArrayIn.GetUpperBound(0) + 1), BytesPerVal * (ArrayIn.GetUpperBound(1) + 1))
    End Function

    '''<summary>Computes the sum of the elements of a vector.</summary>
    '''<param name="ArraySrc">Array to calculate sum from.</param>
    '''<param name="TotalSum">Sum of all elements.</param>
    '''<remarks>Function ippsSum_16s32s_Sfs is calles with scaling factor 0 and sliced.</remarks>
    Public Function Sum(ByRef ArraySrc(,) As Short, ByRef TotalSum As Long) As IppStatus
        Dim RetVal As IppStatus = IppStatus.NoErr
        Dim Caller As System.Delegate = CallIPPS("ippsSum_16s32s_Sfs", GetType(CallSignature.IntPtr_Integer_IntPtr_Integer))
        Using Pinner As New cPinHandler
            TotalSum = 0
            Dim ArrayDst(0) As Integer
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
        End Using : Return RetVal
    End Function


    ' UTILITY FUNCTIONS


    ''' <summary>Build some test data by just counting from 1.</summary>
    ''' <param name="FirstIndexRange">Range of the first index of the returned matrix.</param>
    ''' <param name="SecondIndexRange">Range of the second index of the returned matrix.</param>
    ''' <returns>Matrix with test data.</returns>
    ''' <remarks>The matrix is stored in memory with first iterating over the second index range.</remarks>
    Public Shared Function BuildTestData(ByVal FirstIndexRange As Integer, ByVal SecondIndexRange As Integer) As UInt16(,)
        Dim RetVal(FirstIndexRange - 1, SecondIndexRange - 1) As UInt16
        Dim Counter As UInt16 = 1
        For Idx1 As Integer = 0 To RetVal.GetUpperBound(0)
            For Idx2 As Integer = 0 To RetVal.GetUpperBound(1)
                RetVal(Idx1, Idx2) = Counter : Counter = CUShort(Counter + 1)
            Next Idx2
        Next Idx1
        Return RetVal
    End Function

    Public Shared Sub DisplayArray(ByRef Data() As UInt16, ByVal Size As Integer)
        Dim Out As New List(Of String)
        For Idx1 As Integer = 0 To Data.GetUpperBound(0)
            Out.Add(Data(Idx1).ValRegIndep.PadLeft(Size))
        Next Idx1
        Console.WriteLine(Join(Out.ToArray, "|"))
    End Sub

    Public Shared Sub DisplayArray(ByRef Data(,) As UInt16, ByVal Size As Integer)
        For Idx1 As Integer = 0 To Data.GetUpperBound(0)
            Dim Out As New List(Of String)
            For Idx2 As Integer = 0 To Data.GetUpperBound(1)
                Out.Add(Data(Idx1, Idx2).ValRegIndep.PadLeft(Size))
            Next Idx2
            Console.WriteLine(Join(Out.ToArray, "|"))
        Next Idx1
    End Sub

    '''<summary>Show data as they are stored in memory (use for line-column-problem solving ...).</summary>
    Public Shared Function ShowHowInMemory(ByRef Data(,) As UInt16) As UInt16()
        Dim GACPin As New cIntelIPP.cPinHandler
        Dim Buffer(2 * (Data.Length - 1)) As Byte
        Dim RetVal(Data.Length - 1) As UInt16
        System.Runtime.InteropServices.Marshal.Copy(GACPin.Pin(Data), Buffer, 0, Buffer.Length)
        System.Runtime.InteropServices.Marshal.Copy(Buffer, 0, GACPin.Pin(RetVal), Buffer.Length)
        Return RetVal
    End Function

End Class

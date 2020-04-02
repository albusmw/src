Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

  '''<summary>Predefined numeric formats.</summary>
  Public Enum ePredefFormats
    None
    '''<summary>An integer number.</summary>
    IntegerNumber
    '''<summary>dB.</summary>
    dB
    '''<summary>dBm.</summary>
    dBm
    '''<summary>dBc.</summary>
    dBc
    '''<summary>Hz.</summary>
    Hz
    '''<summary>%.</summary>
    Percentage
    '''<summary>°.</summary>
    Degree
    '''<summary>Seconds.</summary>
    s
    '''<summary>Samples (integer).</summary>
    Samples_Integer
    '''<summary>Samples (floating).</summary>
    Samples_Float
    '''<summary>Volt.</summary>
    Volt
    '''<summary>Volt*Volt</summary>
    VoltVolt
    '''<summary>dBVolt.</summary>
    dBVolt
    '''<summary>Watt.</summary>
    Watt
    '''<summary>Per Second.</summary>
    Per_Second
    '''<summary>Parts per million.</summary>
    ppm
  End Enum

#Region "EnumDescription"

  '''<summary>This converter class can be attached to an enum property and will reflect the Description attribute as drop-down field.</summary>
  Public Class EnumDesciptionConverter : Inherits System.ComponentModel.EnumConverter

    Protected myVal As Type

    Public Sub New(ByVal type As Type)
      MyBase.New(type.GetType)
      myVal = type
    End Sub

    '''<summary>Specify if the property can be expanded by the "+" sign.</summary>
    '''<param name="context"></param>
    '''<returns>Always FALSE as there is nothing to expand here for enums...</returns>
    '''<remarks>Refer to http://www.codeproject.com/KB/cs/propertyeditor.aspx for details.</remarks>
    Public Overrides Function GetPropertiesSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
      Return False
    End Function

    '''<summary>All entries which are displayed in the drop-down list.</summary>
    '''<param name="context">Not of relevance.</param>
    Public Overrides Function GetStandardValues(ByVal context As System.ComponentModel.ITypeDescriptorContext) As StandardValuesCollection
            Dim ComposedValueList As New Collections.ArrayList
            Dim fis As System.Reflection.FieldInfo() = myVal.GetFields()
            For Each fi As System.Reflection.FieldInfo In fis
        Dim attributes As System.ComponentModel.DescriptionAttribute() = CType(fi.GetCustomAttributes(GetType(System.ComponentModel.DescriptionAttribute), False), System.ComponentModel.DescriptionAttribute())
        If attributes.Length > 0 Then ComposedValueList.Add(fi.GetValue(fi.Name))
      Next fi
      Return New StandardValuesCollection(ComposedValueList)
    End Function

    Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object

      'String to enum conversion
      If TypeOf value Is String Then
        Return GetEnumValue(myVal, CStr(value))
      End If

      'Enum to string conversion
      If TypeOf value Is System.Enum Then
        Return GetEnumDescription(CType(value, System.Enum))
      End If

      Return MyBase.ConvertFrom(context, culture, value)

    End Function

    Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object

      'Enum to string conversion
      If (TypeOf value Is System.Enum AndAlso (destinationType Is GetType(String))) Then
        Return GetEnumDescription(DirectCast(value, System.Enum))
      End If

      'Enum to string conversion
      If (TypeOf value Is String AndAlso (destinationType Is GetType(String))) Then
        Return GetEnumDescription(myVal, CStr(value))
      End If

      Return MyBase.ConvertTo(context, culture, value, destinationType)

    End Function

    '================================================================================
    'Helper and service functions (can also be used from outside)
    '================================================================================

    '''<summary>Get the description (derived from the Description attribute) of the passed enum value.</summary>
    '''<param name="value">Value of the enum to read out description.</param>
    '''<returns>Description attribute of this enum or the value.ToString value if the attribute is not present.</returns>
    Public Shared Function GetEnumDescription(ByVal value As System.Enum) As String
      Dim InfoField As System.Reflection.FieldInfo = value.GetType.GetField(value.ToString)
      If IsNothing(InfoField) = False Then
        Dim attributes As System.ComponentModel.DescriptionAttribute() = CType(InfoField.GetCustomAttributes(GetType(System.ComponentModel.DescriptionAttribute), False), System.ComponentModel.DescriptionAttribute())
        If attributes.Length > 0 Then Return attributes(0).Description
      End If
      Return value.ToString
    End Function

    '''<summary>Get the description (derived from the Description attribute) of the passed enum value.</summary>
    '''<param name="value">Type of the enum to read out description.</param>
    '''<param name="name">Name of the enum (as defined by the enum).</param>
    '''<returns>Description attribute of this enum or the value.ToString value if the attribute is not present.</returns>
    Public Shared Function GetEnumDescription(ByVal value As Type, ByVal name As String) As String
      Dim InfoField As System.Reflection.FieldInfo = value.GetField(name)
      If IsNothing(InfoField) = False Then
        Dim attributes As System.ComponentModel.DescriptionAttribute() = CType(InfoField.GetCustomAttributes(GetType(System.ComponentModel.DescriptionAttribute), False), System.ComponentModel.DescriptionAttribute())
        If attributes.Length > 0 Then Return attributes(0).Description
      End If
      Return name
    End Function

    '''<summary>String to enum.</summary>
    Public Shared Function GetEnumValue(ByVal value As Type, ByVal description As String) As Object

      Dim fis As System.Reflection.FieldInfo() = value.GetFields

      For Each fi As System.Reflection.FieldInfo In fis

        Dim attributes As System.ComponentModel.DescriptionAttribute() = CType(fi.GetCustomAttributes(GetType(System.ComponentModel.DescriptionAttribute), False), System.ComponentModel.DescriptionAttribute())

        If ((attributes.Length > 0) AndAlso (attributes(0).Description = description)) Then
          Return fi.GetValue(fi.Name)
        End If

        If (fi.Name = description) Then
          Return fi.GetValue(fi.Name)
        End If

      Next

      Return description

    End Function

  End Class

#End Region

    'This classes can be used as converters for numerical entered values in property grid objects.
    'Formating is done in full precision (indicated by AutoFormat = True) as the value entered by the user is of the user defined precision

    '''<summary>Property converter for unit Hz</summary>
    Public Class DoublePropertyConverter : Inherits System.ComponentModel.DoubleConverter

        Dim PredefFormat As ePredefFormats = ePredefFormats.None

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object) As Object

            'String to double conversion
            If TypeOf value Is String Then Return Helpers.ToDouble(CStr(value), Helpers.PredefFormatString(PredefFormat))

            'Double to string conversion
            If TypeOf value Is Double Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertFrom(context, culture, value)

        End Function

        Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object

            'Double to string conversion
            If (TypeOf value Is System.Double AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(DirectCast(value, System.Double), PredefFormat, True)

            'Double to string conversion
            If (TypeOf value Is String AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertTo(context, culture, value, destinationType)

        End Function

    End Class

    'This classes can be used as converters for numerical entered values in property grid objects.
    'Formating is done in full precision (indicated by AutoFormat = True) as the value entered by the user is of the user defined precision

    '''<summary>Property converter for unit Volt</summary>
    Public Class DoublePropertyConverter_Volt : Inherits System.ComponentModel.DoubleConverter

        Dim PredefFormat As ePredefFormats = ePredefFormats.Volt

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object) As Object

            'String to double conversion
            If TypeOf value Is String Then Return Helpers.ToDouble(CStr(value), Helpers.PredefFormatString(PredefFormat))

            'Double to string conversion
            If TypeOf value Is Double Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertFrom(context, culture, value)

        End Function

        Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object

            'Double to string conversion
            If (TypeOf value Is System.Double AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(DirectCast(value, System.Double), PredefFormat, True)

            'Double to string conversion
            If (TypeOf value Is String AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertTo(context, culture, value, destinationType)

        End Function

    End Class

    '''<summary>Property converter for unit s</summary>
    Public Class DoublePropertyConverter_s : Inherits System.ComponentModel.DoubleConverter

    Dim PredefFormat As ePredefFormats = ePredefFormats.s

    Public Sub New()
      MyBase.New()
    End Sub

    Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object) As Object

      'String to double conversion
      If TypeOf value Is String Then Return Helpers.ToDouble(CStr(value), Helpers.PredefFormatString(PredefFormat))

      'Double to string conversion
      If TypeOf value Is Double Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

      'Default converter
      Return MyBase.ConvertFrom(context, culture, value)

    End Function

    Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object

      'Double to string conversion
      If (TypeOf value Is System.Double AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(DirectCast(value, System.Double), PredefFormat, True)

      'Double to string conversion
      If (TypeOf value Is String AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

      'Default converter
      Return MyBase.ConvertTo(context, culture, value, destinationType)

    End Function

  End Class

    '''<summary>Property converter for unit %</summary>
    Public Class DoublePropertyConverter_Percent : Inherits System.ComponentModel.DoubleConverter

        Dim PredefFormat As ePredefFormats = ePredefFormats.Percentage

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Function ConvertFrom(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object) As Object

            'String to double conversion
            If TypeOf value Is String Then Return Helpers.ToDouble(CStr(value), Helpers.PredefFormatString(PredefFormat))

            'Double to string conversion
            If TypeOf value Is Double Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertFrom(context, culture, value)

        End Function

        Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object

            'Double to string conversion
            If (TypeOf value Is System.Double AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(DirectCast(value, System.Double), PredefFormat, True)

            'Double to string conversion
            If (TypeOf value Is String AndAlso (destinationType Is GetType(String))) Then Return Helpers.Convert(CDbl(value), PredefFormat, True)

            'Default converter
            Return MyBase.ConvertTo(context, culture, value, destinationType)

        End Function

    End Class

    Friend Class Helpers

    '''<summary>Special values common for numerical values.</summary>
    Public Enum eSpecialValue
      None
      MaxValue
      MinValue
    End Enum

    Public Shared Function PredefFormatString(ByVal Predef As ePredefFormats) As String
      Select Case Predef
        Case ePredefFormats.None
          Return String.Empty
        Case ePredefFormats.IntegerNumber
          Return String.Empty
        Case ePredefFormats.dB
          Return "dB"
        Case ePredefFormats.dBm
          Return "dBm"
        Case ePredefFormats.dBc
          Return "dBc"
        Case ePredefFormats.Hz
          Return "Hz"
        Case ePredefFormats.Percentage
          Return "%"
        Case ePredefFormats.Degree
          Return "°"
        Case ePredefFormats.s
          Return "s"
        Case ePredefFormats.Samples_Integer
          Return "Samples"
        Case ePredefFormats.Samples_Float
          Return "Samples"
        Case ePredefFormats.Volt
          Return "V"
        Case ePredefFormats.VoltVolt
          Return "V²"
        Case ePredefFormats.dBVolt
          Return "dBV"
        Case ePredefFormats.Watt
          Return "W"
        Case ePredefFormats.Per_Second
          Return "/s"
        Case ePredefFormats.ppm
          Return "ppm"
        Case Else
          Return String.Empty
      End Select
    End Function

    '''<summary>Get the numeric value of the passed string which can contain E-notation and SI extensions.</summary>
    '''<param name="Value">Value to perform calculation on.</param>
    '''<param name="Unit">Unit (may or may not be used).</param>
    '''<returns>Value of the passed string.</returns>
    Public Shared Function ToDouble(ByVal Value As String, ByVal Unit As String) As Double
      Dim NumericPart As String = String.Empty
      Dim Multiplier As Long, Divider As Long
      Dim SpecialValue As eSpecialValue
      ExtractNumericParts(Value, Unit, NumericPart, Multiplier, Divider, SpecialValue)
      Select Case SpecialValue
        Case eSpecialValue.MaxValue
          Return Double.MaxValue
        Case eSpecialValue.MinValue
          Return Double.MinValue
        Case Else
          If String.IsNullOrEmpty(NumericPart) = False Then
            Dim RetVal As Double = Double.NaN
            If Double.TryParse(StdNumFormat(NumericPart), Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, RetVal) = True Then
              Return RetVal * (Multiplier / Divider)
            End If
          End If
      End Select
      Return Double.NaN
    End Function

    '''<summary>Remove 1000 separators and use only "." as separator.</summary>
    '''<param name="TextToParse">Text to process.</param>
    '''<returns>Text which only contains "." as decimal separator.</returns>
    Public Shared Function StdNumFormat(ByVal TextToParse As String) As String

      If TextToParse.Contains(",") And TextToParse.Contains(".") Then
        If TextToParse.IndexOf(",") < TextToParse.IndexOf(".") Then
          TextToParse = TextToParse.Replace(",", String.Empty)                      '"1,234,567.890" -> "1234567.890"
        Else
          TextToParse = TextToParse.Replace(".", String.Empty).Replace(",", ".")    '"1.234.567,890" -> "1234567.890"
        End If
      Else
        If TextToParse.Contains(",") Then TextToParse = TextToParse.Replace(",", ".")
      End If

      Return TextToParse

    End Function

    '''<summary>Get the relevant numeric parts from the passed value.</summary>
    '''<param name="TextToParse">Value to perform parsing on.</param>
    '''<param name="Unit">Acceptable unit.</param>
    '''<param name="NumberPart">Numeric part, containg only one decimal separator.</param>
    '''<param name="Multiplier">Multiplier extracted.</param>
    Public Shared Sub ExtractNumericParts(ByVal TextToParse As String, ByVal Unit As String, ByRef NumberPart As String, ByRef Multiplier As Long, ByRef Divider As Long, ByRef SpecialValue As eSpecialValue)

      Dim NoNumFormat As System.Globalization.NumberFormatInfo = System.Globalization.NumberFormatInfo.InvariantInfo

      '----------------------------------------
      '0. Init
      Multiplier = 1 : Divider = 1

      '----------------------------------------
      '1. Primary analysis if the string is value

      'If the string is not defined, return 
      If IsNothing(TextToParse) = True Then Exit Sub
      TextToParse = TextToParse.Trim(New Char() {CType(" ", Char), CType(vbCr, Char), CType(vbLf, Char)})
      If IsNothing(TextToParse) = True Then Exit Sub

      '----------------------------------------
      '2. Take care on the units

      'Delete unit ("Hz", "s", ...) if present and remove spaces when present after that
      If TextToParse.EndsWith(Unit) = True Then TextToParse = TextToParse.Substring(0, TextToParse.Length - Unit.Length).Trim

      '----------------------------------------
      '3. Identify exponential markers and cut off

      Multiplier = 1
      If TextToParse.Length > 1 Then
        Select Case TextToParse.Substring(TextToParse.Length - 1, 1)
          Case "k", "K"
            Multiplier = 1000
          Case "M"
            Multiplier = 1000000
          Case "g", "G"
            Multiplier = 1000000000
          Case "T"
            Multiplier = 1000000000000
          Case "m"
            Divider = 1000
          Case "µ", "u", "U"
            Divider = 1000000
          Case "n", "N"
            Divider = 1000000000
          Case "p"
            Divider = 1000000000000
          Case Else
            Multiplier = 1
        End Select
        If Multiplier <> 1 Then TextToParse = TextToParse.Substring(0, TextToParse.Length - 1).Trim
        If Divider <> 1 Then TextToParse = TextToParse.Substring(0, TextToParse.Length - 1).Trim
      End If

      '----------------------------------------
      '4. Force one separator "."
      TextToParse = StdNumFormat(TextToParse)

      '----------------------------------------
      '5. Here, only the number ifself should be present

      NumberPart = String.Empty
      SpecialValue = eSpecialValue.None

      Dim ResultSingle As Single
      If Single.TryParse(TextToParse, Globalization.NumberStyles.Any, NoNumFormat, ResultSingle) = True Then
        NumberPart = TextToParse
      Else
        If TextToParse = Single.MaxValue.ToString(NoNumFormat).Trim Then
          NumberPart = TextToParse : SpecialValue = eSpecialValue.MaxValue
        End If
        If TextToParse = Single.MinValue.ToString(NoNumFormat).Trim Then
          NumberPart = TextToParse : SpecialValue = eSpecialValue.MinValue
        End If
      End If

      Dim ResultDouble As Double
      If Double.TryParse(TextToParse, Globalization.NumberStyles.Any, NoNumFormat, ResultDouble) = True Then
        NumberPart = TextToParse
      Else
        If TextToParse = Double.MaxValue.ToString(NoNumFormat).Trim Then
          NumberPart = TextToParse : SpecialValue = eSpecialValue.MaxValue
        End If
        If TextToParse = Double.MinValue.ToString(NoNumFormat).Trim Then
          NumberPart = TextToParse : SpecialValue = eSpecialValue.MinValue
        End If
      End If
      If String.IsNullOrEmpty(NumberPart) = False Then Exit Sub

      Dim ResultInteger As Integer
      If Integer.TryParse(TextToParse, Globalization.NumberStyles.Any, NoNumFormat, ResultInteger) = True Then
        NumberPart = TextToParse
      End If

      Dim ResultLong As Long
      If Long.TryParse(TextToParse, Globalization.NumberStyles.Any, NoNumFormat, ResultLong) = True Then
        NumberPart = TextToParse
      End If

      '----------------------------------------
      '5. Failed ...
      NumberPart = String.Empty

    End Sub

    Public Shared Function Convert(ByVal Value As Double, ByVal Predef As ePredefFormats, ByVal AutoFormat As Boolean) As String
      Select Case Predef
        Case ePredefFormats.None
          Return Convert(Value, "", "", False, False, True)
        Case ePredefFormats.IntegerNumber
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0")))
        Case ePredefFormats.dB
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.0")))
        Case ePredefFormats.dBm
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.0")))
        Case ePredefFormats.dBc
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.0")))
        Case ePredefFormats.Hz
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Percentage
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Degree
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.s
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Samples_Integer
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0")))
        Case ePredefFormats.Samples_Float
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Volt
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.VoltVolt
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.dBVolt
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Watt
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.Per_Second
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case ePredefFormats.ppm
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
        Case Else
          Return Convert(Value, Predef, CStr(IIf(AutoFormat = True, String.Empty, "0.00")))
      End Select
    End Function

    Public Shared Function Convert(ByVal Value As Double, ByVal Predef As ePredefFormats, ByVal FormatString As String) As String
      Dim Unit As String = PredefFormatString(Predef)
      Select Case Predef
        Case ePredefFormats.None
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.IntegerNumber
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.dB
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.dBm
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.dBc
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.Hz
          If System.Math.Abs(Value) > 1 Then
            Return Convert(Value, Unit, FormatString, True, False, True)
          Else
            Return Convert(Value, Unit, FormatString, False, False, True)
          End If
        Case ePredefFormats.Percentage
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.Degree
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.s
          Return Convert(Value, Unit, FormatString, True, False, True)
        Case ePredefFormats.Samples_Integer
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.Samples_Float
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.Volt
          Return Convert(Value, Unit, FormatString, True, False, True)
        Case ePredefFormats.VoltVolt
          Return Convert(Value, Unit, FormatString, True, False, True)
        Case ePredefFormats.dBVolt
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case ePredefFormats.Watt
          Return Convert(Value, Unit, FormatString, True, False, True)
        Case ePredefFormats.Per_Second
          Return Convert(Value, Unit, FormatString, False, False, True)
        Case Else
          Return Convert(Value, Unit, FormatString, True, False, True)
      End Select
    End Function

    '''<summary>Return a formated string value according SI prefix, ...</summary>
    '''<param name="Value">The value to format.</param>
    '''<param name="BaseUnit">The unit without prefix (e.g. "Hz").</param>
    '''<param name="FormatString">How the value shall be formated.</param>
    '''<param name="UsekMG">Shall kilo, mega, ... be used?</param>
    '''<param name="UsePlus">Shall "+" be added for positiv values.</param>
    '''<param name="RegIndep">Shall only "." as decimal separator be used?</param>
    '''<returns>Formated value as requested.</returns>
    Public Shared Function Convert(ByVal Value As Double, ByVal BaseUnit As String, Optional ByVal FormatString As String = "", Optional ByVal UsekMG As Boolean = True, Optional ByVal UsePlus As Boolean = True, Optional ByVal RegIndep As Boolean = False) As String

      Dim v As Double = Value
      Dim RetVal As String = String.Empty

      RetVal = Format(v, FormatString) & " " & BaseUnit

      'If special format extension should be used
      If UsekMG Then
        If v = 0 Then Convert = "0 " & BaseUnit 'zero
        If System.Math.Abs(v) >= 1 And System.Math.Abs(v) < 1000 Then RetVal = Format(v, FormatString) & " " & BaseUnit 'normal
        If System.Math.Abs(v) >= 1000 And System.Math.Abs(v) < 1000000 Then RetVal = Format(v / 1000, FormatString) & " k" & BaseUnit 'kilo
        If System.Math.Abs(v) >= 1000000 And System.Math.Abs(v) < 1000000000 Then RetVal = Format(v / 1000000, FormatString) & " M" & BaseUnit 'mega
        If System.Math.Abs(v) >= 1000000000 And System.Math.Abs(v) < 1000000000000.0# Then RetVal = Format(v / 1000000000, FormatString) & " G" & BaseUnit 'giga
        If System.Math.Abs(v) >= 0.001 And System.Math.Abs(v) < 1 Then RetVal = Format(v * 1000, FormatString) & " m" & BaseUnit 'milli
        If System.Math.Abs(v) >= 0.00000099999999999999995 And System.Math.Abs(v) < 0.001 Then RetVal = Format(v * 1000000, FormatString) & " µ" & BaseUnit 'mycro
        If System.Math.Abs(v) >= 0.0000000010000000000000001 And System.Math.Abs(v) < 0.00000099999999999999995 Then RetVal = Format(v * 1000000000, FormatString) & " n" & BaseUnit 'nano
        If System.Math.Abs(v) >= 0.00000000000099999999999999998 And System.Math.Abs(v) < 0.0000000010000000000000001 Then RetVal = Format(v * 1000000000000.0#, FormatString) & " p" & BaseUnit 'pico
        If System.Math.Abs(v) >= 0.0000000000000010000000000000001 And System.Math.Abs(v) < 0.00000000000099999999999999998 Then RetVal = Format(v * 1000000000000000.0, FormatString) & " f" & BaseUnit 'pico
      End If

      'Add "+" if required
      If UsePlus = True Then
        If v >= 0 Then RetVal = "+" & RetVal
      End If

      'Only use "."?
      If RegIndep = True Then
        RetVal = StdNumFormat(RetVal)
      End If

      'Trim all
      Return Trim(RetVal)

    End Function

  End Class

End Namespace
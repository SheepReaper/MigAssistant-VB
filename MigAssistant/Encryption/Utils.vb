Imports System.Configuration
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Encryption
    ''' <summary>
    '''     Friend class for shared utility methods used by multiple Encryption classes
    ''' </summary>
    Friend Class Utils
        ''' <summary>
        '''     converts an array of bytes to a string Hex representation
        ''' </summary>
        Friend Shared Function ToHex(ba() As Byte) As String
            If ba Is Nothing OrElse ba.Length = 0 Then
                Return ""
            End If
            Const hexFormat = "{0:X2}"
            Dim sb As New StringBuilder
            For Each b As Byte In ba
                sb.Append(String.Format(hexFormat, b))
            Next
            Return sb.ToString
        End Function

        ''' <summary>
        '''     converts from a string Hex representation to an array of bytes
        ''' </summary>
        Friend Shared Function FromHex(hexEncoded As String) As Byte()
            If hexEncoded Is Nothing OrElse hexEncoded.Length = 0 Then
                Return Nothing
            End If
            Try
                Dim l As Integer = Convert.ToInt32(hexEncoded.Length/2)
                Dim b(l - 1) As Byte
                For i = 0 To l - 1
                    b(i) = Convert.ToByte(hexEncoded.Substring(i*2, 2), 16)
                Next
                Return b
            Catch ex As Exception
                Throw New FormatException("The provided string does not appear to be Hex encoded:" &
                                          Environment.NewLine & hexEncoded & Environment.NewLine, ex)
            End Try
        End Function

        ''' <summary>
        '''     converts from a string Base64 representation to an array of bytes
        ''' </summary>
        Friend Shared Function FromBase64(base64Encoded As String) As Byte()
            If base64Encoded Is Nothing OrElse base64Encoded.Length = 0 Then
                Return Nothing
            End If
            Try
                Return Convert.FromBase64String(base64Encoded)
            Catch ex As FormatException
                Throw New FormatException("The provided string does not appear to be Base64 encoded:" &
                                          Environment.NewLine & base64Encoded & Environment.NewLine, ex)
            End Try
        End Function

        ''' <summary>
        '''     converts from an array of bytes to a string Base64 representation
        ''' </summary>
        Friend Shared Function ToBase64(b() As Byte) As String
            If b Is Nothing OrElse b.Length = 0 Then
                Return ""
            End If
            Return Convert.ToBase64String(b)
        End Function

        ''' <summary>
        '''     retrieve an element from an XML string
        ''' </summary>
        Friend Shared Function GetXmlElement(xml As String, element As String) As String
            Dim m As Match
            m = Regex.Match(xml, "<" & element & ">(?<Element>[^>]*)</" & element & ">", RegexOptions.IgnoreCase)
            If m Is Nothing Then
                Throw New Exception("Could not find <" & element & "></" & element & "> in provided Public Key XML.")
            End If
            Return m.Groups("Element").ToString
        End Function

        ''' <summary>
        '''     Returns the specified string value from the application .config file
        ''' </summary>
        Friend Shared Function GetConfigString(key As String,
                                               Optional ByVal isRequired As Boolean = True) As String

            Dim s = ConfigurationManager.AppSettings.Get(key)
            If s = Nothing Then
                If isRequired Then
                    Throw New ConfigurationErrorsException("key <" & key & "> is missing from .config file")
                Else
                    Return ""
                End If
            Else
                Return s
            End If
        End Function

        Friend Shared Function WriteConfigKey(key As String, value As String) As String
            Dim s As String = "<add key=""{0}"" value=""{1}"" />" & Environment.NewLine
            Return String.Format(s, key, value)
        End Function

        Friend Shared Function WriteXmlElement(element As String, value As String) As String
            Dim s As String = "<{0}>{1}</{0}>" & Environment.NewLine
            Return String.Format(s, element, value)
        End Function

        Friend Shared Function WriteXmlNode(element As String, Optional ByVal isClosing As Boolean = False) As String
            Dim s As String
            If isClosing Then
                s = "</{0}>" & Environment.NewLine
            Else
                s = "<{0}>" & Environment.NewLine
            End If
            Return String.Format(s, element)
        End Function
    End Class
End NameSpace
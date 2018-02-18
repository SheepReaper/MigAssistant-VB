Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.Text

Namespace Encryption
    ''' <summary>
    '''     Asymmetric encryption uses a pair of keys to encrypt and decrypt.
    '''     There is a "public" key which is used to encrypt. Decrypting, on the other hand,
    '''     requires both the "public" key and an additional "private" key. The advantage is
    '''     that people can send you encrypted messages without being able to decrypt them.
    ''' </summary>
    ''' <remarks>
    '''     The only provider supported is the <see cref="RSACryptoServiceProvider" />
    ''' </remarks>
    Public Class Asymmetric
        Private ReadOnly _rsa As RSACryptoServiceProvider
        Private _KeyContainerName As String = "Encryption.AsymmetricEncryption.DefaultContainerName"
        Private _UseMachineKeystore As Boolean = True
        Private ReadOnly _KeySize As Integer = 1024

        Private Const _ElementParent As String = "RSAKeyValue"
        Private Const _ElementModulus As String = "Modulus"
        Private Const _ElementExponent As String = "Exponent"
        Private Const _ElementPrimeP As String = "P"
        Private Const _ElementPrimeQ As String = "Q"
        Private Const _ElementPrimeExponentP As String = "DP"
        Private Const _ElementPrimeExponentQ As String = "DQ"
        Private Const _ElementCoefficient As String = "InverseQ"
        Private Const _ElementPrivateExponent As String = "D"

        '-- http://forum.java.sun.com/thread.jsp?forum=9&thread=552022&tstart=0&trange=15 
        Private Const _KeyModulus As String = "PublicKey.Modulus"
        Private Const _KeyExponent As String = "PublicKey.Exponent"
        Private Const _KeyPrimeP As String = "PrivateKey.P"
        Private Const _KeyPrimeQ As String = "PrivateKey.Q"
        Private Const _KeyPrimeExponentP As String = "PrivateKey.DP"
        Private Const _KeyPrimeExponentQ As String = "PrivateKey.DQ"
        Private Const _KeyCoefficient As String = "PrivateKey.InverseQ"
        Private Const _KeyPrivateExponent As String = "PrivateKey.D"

#Region "  PublicKey Class"

        ''' <summary>
        '''     Represents a public encryption key. Intended to be shared, it
        '''     contains only the Modulus and Exponent.
        ''' </summary>
        Public Class PublicKey
            Public Modulus As String
            Public Exponent As String

            Public Sub New()
            End Sub

            Public Sub New(KeyXml As String)
                LoadFromXml(KeyXml)
            End Sub

            ''' <summary>
            '''     Load public key from App.config or Web.config file
            ''' </summary>
            Public Sub LoadFromConfig()
                Me.Modulus = Utils.GetConfigString(_KeyModulus)
                Me.Exponent = Utils.GetConfigString(_KeyExponent)
            End Sub

            ''' <summary>
            '''     Returns *.config file XML section representing this public key
            ''' </summary>
            Public Function ToConfigSection() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteConfigKey(_KeyModulus, Me.Modulus))
                    .Append(Utils.WriteConfigKey(_KeyExponent, Me.Exponent))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the *.config file representation of this public key to a file
            ''' </summary>
            Public Sub ExportToConfigFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(Me.ToConfigSection)
                sw.Close()
            End Sub

            ''' <summary>
            '''     Loads the public key from its XML string
            ''' </summary>
            Public Sub LoadFromXml(keyXml As String)
                Me.Modulus = Utils.GetXmlElement(keyXml, "Modulus")
                Me.Exponent = Utils.GetXmlElement(keyXml, "Exponent")
            End Sub

            ''' <summary>
            '''     Converts this public key to an RSAParameters object
            ''' </summary>
            Public Function ToParameters() As RSAParameters
                Dim r As New RSAParameters With {
                    .Modulus = Convert.FromBase64String(Me.Modulus),
                    .Exponent = Convert.FromBase64String(Me.Exponent)
                }
                Return r
            End Function

            ''' <summary>
            '''     Converts this public key to its XML string representation
            ''' </summary>
            Public Function ToXml() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteXmlNode(_ElementParent))
                    .Append(Utils.WriteXmlElement(_ElementModulus, Me.Modulus))
                    .Append(Utils.WriteXmlElement(_ElementExponent, Me.Exponent))
                    .Append(Utils.WriteXmlNode(_ElementParent, True))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the Xml representation of this public key to a file
            ''' </summary>
            Public Sub ExportToXmlFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(Me.ToXml)
                sw.Close()
            End Sub
        End Class

#End Region

#Region "  PrivateKey Class"

        ''' <summary>
        '''     Represents a private encryption key. Not intended to be shared, as it
        '''     contains all the elements that make up the key.
        ''' </summary>
        Public Class PrivateKey
            Public Modulus As String
            Public Exponent As String
            Public PrimeP As String
            Public PrimeQ As String
            Public PrimeExponentP As String
            Public PrimeExponentQ As String
            Public Coefficient As String
            Public PrivateExponent As String

            Public Sub New()
            End Sub

            Public Sub New(keyXml As String)
                LoadFromXml(keyXml)
            End Sub

            ''' <summary>
            '''     Load private key from App.config or Web.config file
            ''' </summary>
            Public Sub LoadFromConfig()
                Me.Modulus = Utils.GetConfigString(_KeyModulus)
                Me.Exponent = Utils.GetConfigString(_KeyExponent)
                Me.PrimeP = Utils.GetConfigString(_KeyPrimeP)
                Me.PrimeQ = Utils.GetConfigString(_KeyPrimeQ)
                Me.PrimeExponentP = Utils.GetConfigString(_KeyPrimeExponentP)
                Me.PrimeExponentQ = Utils.GetConfigString(_KeyPrimeExponentQ)
                Me.Coefficient = Utils.GetConfigString(_KeyCoefficient)
                Me.PrivateExponent = Utils.GetConfigString(_KeyPrivateExponent)
            End Sub

            ''' <summary>
            '''     Converts this private key to an RSAParameters object
            ''' </summary>
            Public Function ToParameters() As RSAParameters
                Dim r As New RSAParameters With {
                    .Modulus = Convert.FromBase64String(Me.Modulus),
                    .Exponent = Convert.FromBase64String(Me.Exponent),
                    .P = Convert.FromBase64String(Me.PrimeP),
                    .Q = Convert.FromBase64String(Me.PrimeQ),
                    .DP = Convert.FromBase64String(Me.PrimeExponentP),
                    .DQ = Convert.FromBase64String(Me.PrimeExponentQ),
                    .InverseQ = Convert.FromBase64String(Me.Coefficient),
                    .D = Convert.FromBase64String(Me.PrivateExponent)
                }
                Return r
            End Function

            ''' <summary>
            '''     Returns *.config file XML section representing this private key
            ''' </summary>
            Public Function ToConfigSection() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteConfigKey(_KeyModulus, Me.Modulus))
                    .Append(Utils.WriteConfigKey(_KeyExponent, Me.Exponent))
                    .Append(Utils.WriteConfigKey(_KeyPrimeP, Me.PrimeP))
                    .Append(Utils.WriteConfigKey(_KeyPrimeQ, Me.PrimeQ))
                    .Append(Utils.WriteConfigKey(_KeyPrimeExponentP, Me.PrimeExponentP))
                    .Append(Utils.WriteConfigKey(_KeyPrimeExponentQ, Me.PrimeExponentQ))
                    .Append(Utils.WriteConfigKey(_KeyCoefficient, Me.Coefficient))
                    .Append(Utils.WriteConfigKey(_KeyPrivateExponent, Me.PrivateExponent))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the *.config file representation of this private key to a file
            ''' </summary>
            Public Sub ExportToConfigFile(strFilePath As String)
                Dim sw As New StreamWriter(strFilePath, False)
                sw.Write(Me.ToConfigSection)
                sw.Close()
            End Sub

            ''' <summary>
            '''     Loads the private key from its XML string
            ''' </summary>
            Public Sub LoadFromXml(keyXml As String)
                Me.Modulus = Utils.GetXmlElement(keyXml, "Modulus")
                Me.Exponent = Utils.GetXmlElement(keyXml, "Exponent")
                Me.PrimeP = Utils.GetXmlElement(keyXml, "P")
                Me.PrimeQ = Utils.GetXmlElement(keyXml, "Q")
                Me.PrimeExponentP = Utils.GetXmlElement(keyXml, "DP")
                Me.PrimeExponentQ = Utils.GetXmlElement(keyXml, "DQ")
                Me.Coefficient = Utils.GetXmlElement(keyXml, "InverseQ")
                Me.PrivateExponent = Utils.GetXmlElement(keyXml, "D")
            End Sub

            ''' <summary>
            '''     Converts this private key to its XML string representation
            ''' </summary>
            Public Function ToXml() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteXmlNode(_ElementParent))
                    .Append(Utils.WriteXmlElement(_ElementModulus, Me.Modulus))
                    .Append(Utils.WriteXmlElement(_ElementExponent, Me.Exponent))
                    .Append(Utils.WriteXmlElement(_ElementPrimeP, Me.PrimeP))
                    .Append(Utils.WriteXmlElement(_ElementPrimeQ, Me.PrimeQ))
                    .Append(Utils.WriteXmlElement(_ElementPrimeExponentP, Me.PrimeExponentP))
                    .Append(Utils.WriteXmlElement(_ElementPrimeExponentQ, Me.PrimeExponentQ))
                    .Append(Utils.WriteXmlElement(_ElementCoefficient, Me.Coefficient))
                    .Append(Utils.WriteXmlElement(_ElementPrivateExponent, Me.PrivateExponent))
                    .Append(Utils.WriteXmlNode(_ElementParent, True))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the Xml representation of this private key to a file
            ''' </summary>
            Public Sub ExportToXmlFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(Me.ToXml)
                sw.Close()
            End Sub
        End Class

#End Region

        ''' <summary>
        '''     Instantiates a new asymmetric encryption session using the default key size;
        '''     this is usally 1024 bits
        ''' </summary>
        Public Sub New()
            _rsa = GetRSAProvider()
        End Sub

        ''' <summary>
        '''     Instantiates a new asymmetric encryption session using a specific key size
        ''' </summary>
        Public Sub New(keySize As Integer)
            _KeySize = keySize
            _rsa = GetRSAProvider()
        End Sub

        ''' <summary>
        '''     Sets the name of the key container used to store this key on disk; this is an
        '''     unavoidable side effect of the underlying Microsoft CryptoAPI.
        ''' </summary>
        ''' <remarks>
        '''     http://support.microsoft.com/default.aspx?scid=http://support.microsoft.com:80/support/kb/articles/q322/3/71.asp
        '''     &amp;NoWebContent=1
        ''' </remarks>
        Public Property KeyContainerName As String
            Get
                Return _KeyContainerName
            End Get
            Set
                _KeyContainerName = Value
            End Set
        End Property

        ''' <summary>
        '''     Returns the current key size, in bits
        ''' </summary>
        Public ReadOnly Property KeySizeBits As Integer
            Get
                Return _rsa.KeySize
            End Get
        End Property

        ''' <summary>
        '''     Returns the maximum supported key size, in bits
        ''' </summary>
        Public ReadOnly Property KeySizeMaxBits As Integer
            Get
                Return _rsa.LegalKeySizes(0).MaxSize
            End Get
        End Property

        ''' <summary>
        '''     Returns the minimum supported key size, in bits
        ''' </summary>
        Public ReadOnly Property KeySizeMinBits As Integer
            Get
                Return _rsa.LegalKeySizes(0).MinSize
            End Get
        End Property

        ''' <summary>
        '''     Returns valid key step sizes, in bits
        ''' </summary>
        Public ReadOnly Property KeySizeStepBits As Integer
            Get
                Return _rsa.LegalKeySizes(0).SkipSize
            End Get
        End Property

        ''' <summary>
        '''     Returns the default public key as stored in the *.config file
        ''' </summary>
        Public ReadOnly Property DefaultPublicKey As PublicKey
            Get
                Dim pubkey As New PublicKey
                pubkey.LoadFromConfig()
                Return pubkey
            End Get
        End Property

        ''' <summary>
        '''     Returns the default private key as stored in the *.config file
        ''' </summary>
        Public ReadOnly Property DefaultPrivateKey As PrivateKey
            Get
                Dim privkey As New PrivateKey
                privkey.LoadFromConfig()
                Return privkey
            End Get
        End Property

        ''' <summary>
        '''     Generates a new public/private key pair as objects
        ''' </summary>
        Public Sub GenerateNewKeyset(ByRef publicKey As PublicKey, ByRef privateKey As PrivateKey)
            Dim PublicKeyXML As String = Nothing
            Dim PrivateKeyXML As String = Nothing
            GenerateNewKeyset(PublicKeyXML, PrivateKeyXML)
            publicKey = New PublicKey(PublicKeyXML)
            privateKey = New PrivateKey(PrivateKeyXML)
        End Sub

        ''' <summary>
        '''     Generates a new public/private key pair as XML strings
        ''' </summary>
        Public Sub GenerateNewKeyset(ByRef publicKeyXML As String, ByRef privateKeyXML As String)
            Dim rsa As RSA = RSACryptoServiceProvider.Create
            publicKeyXML = rsa.ToXmlString(False)
            privateKeyXML = rsa.ToXmlString(True)
        End Sub

        ''' <summary>
        '''     Encrypts data using the default public key
        ''' </summary>
        Public Function Encrypt(d As Data) As Data
            Dim PublicKey As PublicKey = DefaultPublicKey
            Return Encrypt(d, PublicKey)
        End Function

        ''' <summary>
        '''     Encrypts data using the provided public key
        ''' </summary>
        Public Function Encrypt(d As Data, publicKey As PublicKey) As Data
            _rsa.ImportParameters(publicKey.ToParameters)
            Return EncryptPrivate(d)
        End Function

        ''' <summary>
        '''     Encrypts data using the provided public key as XML
        ''' </summary>
        Public Function Encrypt(d As Data, publicKeyXML As String) As Data
            LoadKeyXml(publicKeyXML, False)
            Return EncryptPrivate(d)
        End Function

        Private Function EncryptPrivate(d As Data) As Data
            Try
                Return New Data(_rsa.Encrypt(d.Bytes, False))
            Catch ex As CryptographicException
                If ex.Message.ToLower.IndexOf("bad length") > - 1 Then
                    Throw _
                        New CryptographicException(
                            "Your data is too large; RSA encryption is designed to encrypt relatively small amounts of data. The exact byte limit depends on the key size. To encrypt more data, use symmetric encryption and then encrypt that symmetric key with asymmetric RSA encryption.",
                            ex)
                Else
                    Throw
                End If
            End Try
        End Function

        ''' <summary>
        '''     Decrypts data using the default private key
        ''' </summary>
        Public Function Decrypt(encryptedData As Data) As Data
            Dim PrivateKey As New PrivateKey
            PrivateKey.LoadFromConfig()
            Return Decrypt(encryptedData, PrivateKey)
        End Function

        ''' <summary>
        '''     Decrypts data using the provided private key
        ''' </summary>
        Public Function Decrypt(encryptedData As Data, PrivateKey As PrivateKey) As Data
            _rsa.ImportParameters(PrivateKey.ToParameters)
            Return DecryptPrivate(encryptedData)
        End Function

        ''' <summary>
        '''     Decrypts data using the provided private key as XML
        ''' </summary>
        Public Function Decrypt(encryptedData As Data, PrivateKeyXML As String) As Data
            LoadKeyXml(PrivateKeyXML, True)
            Return DecryptPrivate(encryptedData)
        End Function

        Private Sub LoadKeyXml(keyXml As String, isPrivate As Boolean)
            Try
                _rsa.FromXmlString(keyXml)
            Catch ex As XmlSyntaxException
                Dim s As String
                If isPrivate Then
                    s = "private"
                Else
                    s = "public"
                End If
                Throw New XmlSyntaxException(
                    $"The provided {s } encryption key XML does not appear to be valid.", ex)
            End Try
        End Sub

        Private Function DecryptPrivate(encryptedData As Data) As Data
            Return New Data(_rsa.Decrypt(encryptedData.Bytes, False))
        End Function

        ''' <summary>
        '''     gets the default RSA provider using the specified key size;
        '''     note that Microsoft's CryptoAPI has an underlying file system dependency that is unavoidable
        ''' </summary>
        ''' <remarks>
        '''     http://support.microsoft.com/default.aspx?scid=http://support.microsoft.com:80/support/kb/articles/q322/3/71.asp
        '''     &amp;NoWebContent=1
        ''' </remarks>
        Private Function GetRSAProvider() As RSACryptoServiceProvider
            Dim rsa As RSACryptoServiceProvider = Nothing
            Dim csp As CspParameters = Nothing
            Try
                csp = New CspParameters With {
                    .KeyContainerName = _KeyContainerName
                }
                rsa = New RSACryptoServiceProvider(_KeySize, csp) With {
                    .PersistKeyInCsp = False
                }
                RSACryptoServiceProvider.UseMachineKeyStore = True
                Return rsa
            Catch ex As CryptographicException
                If ex.Message.ToLower.IndexOf("csp for this implementation could not be acquired") > - 1 Then
                    Throw New Exception("Unable to obtain Cryptographic Service Provider. " &
                                        "Either the permissions are incorrect on the " &
                                        "'C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA\MachineKeys' " &
                                        "folder, or the current security context '" & WindowsIdentity.GetCurrent.Name &
                                        "'" &
                                        " does not have access to this folder.", ex)
                Else
                    Throw
                End If
            Finally
                If Not rsa Is Nothing Then
                    rsa = Nothing
                End If
                If Not csp Is Nothing Then
                    csp = Nothing
                End If
            End Try
        End Function
    End Class
End NameSpace
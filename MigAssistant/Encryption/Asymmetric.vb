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
    '''     The only provider supported is the <see cref="Security.Cryptography.RSACryptoServiceProvider" />
    ''' </remarks>
    Public Class Asymmetric
        Private ReadOnly _rsa As RSACryptoServiceProvider
        Private _keyContainerName As String = "Encryption.AsymmetricEncryption.DefaultContainerName"
        'Private _useMachineKeystore As Boolean = True
        Private ReadOnly _keySize As Integer = 1024

        Private Const ElementParent As String = "RSAKeyValue"
        Private Const ElementModulus As String = "Modulus"
        Private Const ElementExponent As String = "Exponent"
        Private Const ElementPrimeP As String = "P"
        Private Const ElementPrimeQ As String = "Q"
        Private Const ElementPrimeExponentP As String = "DP"
        Private Const ElementPrimeExponentQ As String = "DQ"
        Private Const ElementCoefficient As String = "InverseQ"
        Private Const ElementPrivateExponent As String = "D"

        '-- http://forum.java.sun.com/thread.jsp?forum=9&thread=552022&tstart=0&trange=15 
        Private Const KeyModulus As String = "PublicKey.Modulus"
        Private Const KeyExponent As String = "PublicKey.Exponent"
        Private Const KeyPrimeP As String = "PrivateKey.P"
        Private Const KeyPrimeQ As String = "PrivateKey.Q"
        Private Const KeyPrimeExponentP As String = "PrivateKey.DP"
        Private Const KeyPrimeExponentQ As String = "PrivateKey.DQ"
        Private Const KeyCoefficient As String = "PrivateKey.InverseQ"
        Private Const KeyPrivateExponent As String = "PrivateKey.D"

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

            Public Sub New(keyXml As String)
                LoadFromXml(keyXml)
            End Sub

            ''' <summary>
            '''     Load public key from App.config or Web.config file
            ''' </summary>
            Public Sub LoadFromConfig()
                Modulus = Utils.GetConfigString(KeyModulus)
                Exponent = Utils.GetConfigString(KeyExponent)
            End Sub

            ''' <summary>
            '''     Returns *.config file XML section representing this public key
            ''' </summary>
            Public Function ToConfigSection() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteConfigKey(KeyModulus, Modulus))
                    .Append(Utils.WriteConfigKey(KeyExponent, Exponent))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the *.config file representation of this public key to a file
            ''' </summary>
            Public Sub ExportToConfigFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(ToConfigSection)
                sw.Close()
            End Sub

            ''' <summary>
            '''     Loads the public key from its XML string
            ''' </summary>
            Public Sub LoadFromXml(keyXml As String)
                Modulus = Utils.GetXmlElement(keyXml, "Modulus")
                Exponent = Utils.GetXmlElement(keyXml, "Exponent")
            End Sub

            ''' <summary>
            '''     Converts this public key to an RSAParameters object
            ''' </summary>
            Public Function ToParameters() As RSAParameters
                Dim r As New RSAParameters With {
                        .Modulus = Convert.FromBase64String(Modulus),
                        .Exponent = Convert.FromBase64String(Exponent)
                        }
                Return r
            End Function

            ''' <summary>
            '''     Converts this public key to its XML string representation
            ''' </summary>
            Public Function ToXml() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteXmlNode(ElementParent))
                    .Append(Utils.WriteXmlElement(ElementModulus, Modulus))
                    .Append(Utils.WriteXmlElement(ElementExponent, Exponent))
                    .Append(Utils.WriteXmlNode(ElementParent, True))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the Xml representation of this public key to a file
            ''' </summary>
            Public Sub ExportToXmlFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(ToXml)
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
                Modulus = Utils.GetConfigString(KeyModulus)
                Exponent = Utils.GetConfigString(KeyExponent)
                PrimeP = Utils.GetConfigString(KeyPrimeP)
                PrimeQ = Utils.GetConfigString(KeyPrimeQ)
                PrimeExponentP = Utils.GetConfigString(KeyPrimeExponentP)
                PrimeExponentQ = Utils.GetConfigString(KeyPrimeExponentQ)
                Coefficient = Utils.GetConfigString(KeyCoefficient)
                PrivateExponent = Utils.GetConfigString(KeyPrivateExponent)
            End Sub

            ''' <summary>
            '''     Converts this private key to an RSAParameters object
            ''' </summary>
            Public Function ToParameters() As RSAParameters
                Dim r As New RSAParameters With {
                        .Modulus = Convert.FromBase64String(Modulus),
                        .Exponent = Convert.FromBase64String(Exponent),
                        .P = Convert.FromBase64String(PrimeP),
                        .Q = Convert.FromBase64String(PrimeQ),
                        .DP = Convert.FromBase64String(PrimeExponentP),
                        .DQ = Convert.FromBase64String(PrimeExponentQ),
                        .InverseQ = Convert.FromBase64String(Coefficient),
                        .D = Convert.FromBase64String(PrivateExponent)
                        }
                Return r
            End Function

            ''' <summary>
            '''     Returns *.config file XML section representing this private key
            ''' </summary>
            Public Function ToConfigSection() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteConfigKey(KeyModulus, Modulus))
                    .Append(Utils.WriteConfigKey(KeyExponent, Exponent))
                    .Append(Utils.WriteConfigKey(KeyPrimeP, PrimeP))
                    .Append(Utils.WriteConfigKey(KeyPrimeQ, PrimeQ))
                    .Append(Utils.WriteConfigKey(KeyPrimeExponentP, PrimeExponentP))
                    .Append(Utils.WriteConfigKey(KeyPrimeExponentQ, PrimeExponentQ))
                    .Append(Utils.WriteConfigKey(KeyCoefficient, Coefficient))
                    .Append(Utils.WriteConfigKey(KeyPrivateExponent, PrivateExponent))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the *.config file representation of this private key to a file
            ''' </summary>
            Public Sub ExportToConfigFile(strFilePath As String)
                Dim sw As New StreamWriter(strFilePath, False)
                sw.Write(ToConfigSection)
                sw.Close()
            End Sub

            ''' <summary>
            '''     Loads the private key from its XML string
            ''' </summary>
            Public Sub LoadFromXml(keyXml As String)
                Modulus = Utils.GetXmlElement(keyXml, "Modulus")
                Exponent = Utils.GetXmlElement(keyXml, "Exponent")
                PrimeP = Utils.GetXmlElement(keyXml, "P")
                PrimeQ = Utils.GetXmlElement(keyXml, "Q")
                PrimeExponentP = Utils.GetXmlElement(keyXml, "DP")
                PrimeExponentQ = Utils.GetXmlElement(keyXml, "DQ")
                Coefficient = Utils.GetXmlElement(keyXml, "InverseQ")
                PrivateExponent = Utils.GetXmlElement(keyXml, "D")
            End Sub

            ''' <summary>
            '''     Converts this private key to its XML string representation
            ''' </summary>
            Public Function ToXml() As String
                Dim sb As New StringBuilder
                With sb
                    .Append(Utils.WriteXmlNode(ElementParent))
                    .Append(Utils.WriteXmlElement(ElementModulus, Modulus))
                    .Append(Utils.WriteXmlElement(ElementExponent, Exponent))
                    .Append(Utils.WriteXmlElement(ElementPrimeP, PrimeP))
                    .Append(Utils.WriteXmlElement(ElementPrimeQ, PrimeQ))
                    .Append(Utils.WriteXmlElement(ElementPrimeExponentP, PrimeExponentP))
                    .Append(Utils.WriteXmlElement(ElementPrimeExponentQ, PrimeExponentQ))
                    .Append(Utils.WriteXmlElement(ElementCoefficient, Coefficient))
                    .Append(Utils.WriteXmlElement(ElementPrivateExponent, PrivateExponent))
                    .Append(Utils.WriteXmlNode(ElementParent, True))
                End With
                Return sb.ToString
            End Function

            ''' <summary>
            '''     Writes the Xml representation of this private key to a file
            ''' </summary>
            Public Sub ExportToXmlFile(filePath As String)
                Dim sw As New StreamWriter(filePath, False)
                sw.Write(ToXml)
                sw.Close()
            End Sub
        End Class

#End Region

        ''' <summary>
        '''     Instantiates a new asymmetric encryption session using the default key size;
        '''     this is usally 1024 bits
        ''' </summary>
        Public Sub New()
            _rsa = GetRsaProvider()
        End Sub

        ''' <summary>
        '''     Instantiates a new asymmetric encryption session using a specific key size
        ''' </summary>
        Public Sub New(keySize As Integer)
            _keySize = keySize
            _rsa = GetRsaProvider()
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
                Return _keyContainerName
            End Get
            Set
                _keyContainerName = Value
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
            Dim publicKeyXml As String = Nothing
            Dim privateKeyXml As String = Nothing
            GenerateNewKeyset(publicKeyXml, privateKeyXml)
            publicKey = New PublicKey(publicKeyXml)
            privateKey = New PrivateKey(privateKeyXml)
        End Sub

        ''' <summary>
        '''     Generates a new public/private key pair as XML strings
        ''' </summary>
        Public Sub GenerateNewKeyset(ByRef publicKeyXml As String, ByRef privateKeyXml As String)
            Dim rsa As RSA = RSACryptoServiceProvider.Create
            publicKeyXml = rsa.ToXmlString(False)
            privateKeyXml = rsa.ToXmlString(True)
        End Sub

        ''' <summary>
        '''     Encrypts data using the default public key
        ''' </summary>
        Public Function Encrypt(d As Data) As Data
            Dim publicKey As PublicKey = DefaultPublicKey
            Return Encrypt(d, publicKey)
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
        Public Function Encrypt(d As Data, publicKeyXml As String) As Data
            LoadKeyXml(publicKeyXml, False)
            Return EncryptPrivate(d)
        End Function

        Private Function EncryptPrivate(d As Data) As Data
            Try
                Return New Data(_rsa.Encrypt(d.Bytes, False))
            Catch ex As CryptographicException
                If ex.Message.ToLower.IndexOf("bad length", StringComparison.Ordinal) > - 1 Then
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
            Dim privateKey As New PrivateKey
            privateKey.LoadFromConfig()
            Return Decrypt(encryptedData, privateKey)
        End Function

        ''' <summary>
        '''     Decrypts data using the provided private key
        ''' </summary>
        Public Function Decrypt(encryptedData As Data, privateKey As PrivateKey) As Data
            _rsa.ImportParameters(privateKey.ToParameters)
            Return DecryptPrivate(encryptedData)
        End Function

        ''' <summary>
        '''     Decrypts data using the provided private key as XML
        ''' </summary>
        Public Function Decrypt(encryptedData As Data, privateKeyXml As String) As Data
            LoadKeyXml(privateKeyXml, True)
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
        Private Function GetRsaProvider() As RSACryptoServiceProvider
            Dim rsa As RSACryptoServiceProvider = Nothing
            Dim csp As CspParameters
            Try
                csp = New CspParameters With {
                    .KeyContainerName = _keyContainerName
                    }
                rsa = New RSACryptoServiceProvider(_keySize, csp) With {
                    .PersistKeyInCsp = False
                    }
                RSACryptoServiceProvider.UseMachineKeyStore = True
                Return rsa
            Catch ex As CryptographicException
                If _
                    ex.Message.ToLower.IndexOf("csp for this implementation could not be acquired",
                                               StringComparison.Ordinal) > - 1 Then
                    Throw _
                        New Exception(
                            $"Unable to obtain Cryptographic Service Provider. Either the permissions are incorrect on the 'C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA\MachineKeys' folder, or the current security context '{ _
                                         WindowsIdentity.GetCurrent.Name}' does not have access to this folder.", ex)
                Else
                    Throw
                End If
            Finally
                If Not rsa Is Nothing Then
                    rsa.Clear()
                End If
                'If Not csp Is Nothing Then
                '    csp = Nothing
                'End If
            End Try
        End Function
    End Class
End NameSpace
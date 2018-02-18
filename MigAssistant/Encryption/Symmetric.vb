Imports System.IO
Imports System.Security.Cryptography

Namespace Encryption
    ''' <summary>
    '''     Symmetric encryption uses a single key to encrypt and decrypt.
    '''     Both parties (encryptor and decryptor) must share the same secret key.
    ''' </summary>
    Public Class Symmetric
        Private Const _DefaultIntializationVector As String = "%1Az=-@qT"
        Private Const _BufferSize As Integer = 2048

        Public Enum Provider
            ''' <summary>
            '''     The Data Encryption Standard provider supports a 64 bit key only
            ''' </summary>
            DES
            ''' <summary>
            '''     The Rivest Cipher 2 provider supports keys ranging from 40 to 128 bits, default is 128 bits
            ''' </summary>
            RC2
            ''' <summary>
            '''     The Rijndael (also known as AES) provider supports keys of 128, 192, or 256 bits with a default of 256 bits
            ''' </summary>
            Rijndael
            ''' <summary>
            '''     The TripleDES provider (also known as 3DES) supports keys of 128 or 192 bits with a default of 192 bits
            ''' </summary>
            TripleDES
        End Enum

        Private _data As Data
        Private _key As Data
        Private _iv As Data
        Private ReadOnly _crypto As SymmetricAlgorithm
        Private _EncryptedBytes As Byte()
        Private _UseDefaultInitializationVector As Boolean

        Private Sub New()
        End Sub

        ''' <summary>
        '''     Instantiates a new symmetric encryption object using the specified provider.
        ''' </summary>
        Public Sub New(provider As Provider, Optional ByVal useDefaultInitializationVector As Boolean = True)
            Select Case provider
                Case provider.DES
                    _crypto = New DESCryptoServiceProvider
                Case provider.RC2
                    _crypto = New RC2CryptoServiceProvider
                Case provider.Rijndael
                    _crypto = New RijndaelManaged
                Case provider.TripleDES
                    _crypto = New TripleDESCryptoServiceProvider
            End Select

            '-- make sure key and IV are always set, no matter what
            Me.Key = RandomKey()
            If useDefaultInitializationVector Then
                Me.IntializationVector = New Data(_DefaultIntializationVector)
            Else
                Me.IntializationVector = RandomInitializationVector()
            End If
        End Sub

        ''' <summary>
        '''     Key size in bytes. We use the default key size for any given provider; if you
        '''     want to force a specific key size, set this property
        ''' </summary>
        Public Property KeySizeBytes As Integer
            Get
                Return _crypto.KeySize\8
            End Get
            Set
                _crypto.KeySize = Value*8
                _key.MaxBytes = Value
            End Set
        End Property

        ''' <summary>
        '''     Key size in bits. We use the default key size for any given provider; if you
        '''     want to force a specific key size, set this property
        ''' </summary>
        Public Property KeySizeBits As Integer
            Get
                Return _crypto.KeySize
            End Get
            Set
                _crypto.KeySize = Value
                _key.MaxBits = Value
            End Set
        End Property

        ''' <summary>
        '''     The key used to encrypt/decrypt data
        ''' </summary>
        Public Property Key As Data
            Get
                Return _key
            End Get
            Set
                _key = Value
                _key.MaxBytes = _crypto.LegalKeySizes(0).MaxSize\8
                _key.MinBytes = _crypto.LegalKeySizes(0).MinSize\8
                _key.StepBytes = _crypto.LegalKeySizes(0).SkipSize\8
            End Set
        End Property

        ''' <summary>
        '''     Using the default Cipher Block Chaining (CBC) mode, all data blocks are processed using
        '''     the value derived from the previous block; the first data block has no previous data block
        '''     to use, so it needs an InitializationVector to feed the first block
        ''' </summary>
        Public Property IntializationVector As Data
            Get
                Return _iv
            End Get
            Set
                _iv = Value
                _iv.MaxBytes = _crypto.BlockSize\8
                _iv.MinBytes = _crypto.BlockSize\8
            End Set
        End Property

        ''' <summary>
        '''     generates a random Initialization Vector, if one was not provided
        ''' </summary>
        Public Function RandomInitializationVector() As Data
            _crypto.GenerateIV()
            Dim d As New Data(_crypto.IV)
            Return d
        End Function

        ''' <summary>
        '''     generates a random Key, if one was not provided
        ''' </summary>
        Public Function RandomKey() As Data
            _crypto.GenerateKey()
            Dim d As New Data(_crypto.Key)
            Return d
        End Function

        ''' <summary>
        '''     Ensures that _crypto object has valid Key and IV
        '''     prior to any attempt to encrypt/decrypt anything
        ''' </summary>
        Private Sub ValidateKeyAndIv(isEncrypting As Boolean)
            If _key.IsEmpty Then
                If isEncrypting Then
                    _key = RandomKey()
                Else
                    Throw New CryptographicException("No key was provided for the decryption operation!")
                End If
            End If
            If _iv.IsEmpty Then
                If isEncrypting Then
                    _iv = RandomInitializationVector()
                Else
                    Throw _
                        New CryptographicException("No initialization vector was provided for the decryption operation!")
                End If
            End If
            _crypto.Key = _key.Bytes
            _crypto.IV = _iv.Bytes
        End Sub

        ''' <summary>
        '''     Encrypts the specified Data using provided key
        ''' </summary>
        Public Function Encrypt(d As Data, key As Data) As Data
            Me.Key = key
            Return Encrypt(d)
        End Function

        ''' <summary>
        '''     Encrypts the specified Data using preset key and preset initialization vector
        ''' </summary>
        Public Function Encrypt(d As Data) As Data
            Dim ms As New MemoryStream

            ValidateKeyAndIv(True)

            Dim cs As New CryptoStream(ms, _crypto.CreateEncryptor(), CryptoStreamMode.Write)
            cs.Write(d.Bytes, 0, d.Bytes.Length)
            cs.Close()
            ms.Close()

            Return New Data(ms.ToArray)
        End Function

        ''' <summary>
        '''     Encrypts the stream to memory using provided key and provided initialization vector
        ''' </summary>
        Public Function Encrypt(s As Stream, key As Data, iv As Data) As Data
            Me.IntializationVector = iv
            Me.Key = key
            Return Encrypt(s)
        End Function

        ''' <summary>
        '''     Encrypts the stream to memory using specified key
        ''' </summary>
        Public Function Encrypt(s As Stream, key As Data) As Data
            Me.Key = key
            Return Encrypt(s)
        End Function

        ''' <summary>
        '''     Encrypts the specified stream to memory using preset key and preset initialization vector
        ''' </summary>
        Public Function Encrypt(s As Stream) As Data
            Dim ms As New MemoryStream
            Dim b(_BufferSize) As Byte
            Dim i As Integer

            ValidateKeyAndIv(True)

            Dim cs As New CryptoStream(ms, _crypto.CreateEncryptor(), CryptoStreamMode.Write)
            i = s.Read(b, 0, _BufferSize)
            Do While i > 0
                cs.Write(b, 0, i)
                i = s.Read(b, 0, _BufferSize)
            Loop

            cs.Close()
            ms.Close()

            Return New Data(ms.ToArray)
        End Function

        ''' <summary>
        '''     Decrypts the specified data using provided key and preset initialization vector
        ''' </summary>
        Public Function Decrypt(encryptedData As Data, key As Data) As Data
            Me.Key = key
            Return Decrypt(encryptedData)
        End Function

        ''' <summary>
        '''     Decrypts the specified stream using provided key and preset initialization vector
        ''' </summary>
        Public Function Decrypt(encryptedStream As Stream, key As Data) As Data
            Me.Key = key
            Return Decrypt(encryptedStream)
        End Function

        ''' <summary>
        '''     Decrypts the specified stream using preset key and preset initialization vector
        ''' </summary>
        Public Function Decrypt(encryptedStream As Stream) As Data
            Dim ms As New MemoryStream
            Dim b(_BufferSize) As Byte

            ValidateKeyAndIv(False)
            Dim cs As New CryptoStream(encryptedStream,
                                       _crypto.CreateDecryptor(), CryptoStreamMode.Read)

            Dim i As Integer
            i = cs.Read(b, 0, _BufferSize)

            Do While i > 0
                ms.Write(b, 0, i)
                i = cs.Read(b, 0, _BufferSize)
            Loop
            cs.Close()
            ms.Close()

            Return New Data(ms.ToArray)
        End Function

        ''' <summary>
        '''     Decrypts the specified data using preset key and preset initialization vector
        ''' </summary>
        Public Function Decrypt(encryptedData As Data) As Data
            Dim ms As New MemoryStream(encryptedData.Bytes, 0, encryptedData.Bytes.Length)
            Dim b = New Byte(encryptedData.Bytes.Length - 1) {}

            ValidateKeyAndIv(False)
            Dim cs As New CryptoStream(ms, _crypto.CreateDecryptor(), CryptoStreamMode.Read)

            Try
                cs.Read(b, 0, encryptedData.Bytes.Length - 1)
            Catch ex As CryptographicException
                Throw New CryptographicException("Unable to decrypt data. The provided key may be invalid.", ex)
            Finally
                cs.Close()
            End Try
            Return New Data(b)
        End Function
    End Class
End NameSpace
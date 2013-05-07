Attribute VB_Name = "modEncryption"
Public Function CryptData(ByVal Str As String, ByVal Password As String) As String
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=231798[/url]
    '
    Dim SK As Long, K As Long

    ' init randomizer for password
    Rnd -1
    Randomize Len(Password)
    ' (((K Mod 256) Xor Asc(Mid$(Password, K, 1))) Xor Fix(256 * Rnd)) -> makes sure that a
    ' password like "pass12" does NOT give the same result as the password "sspa12" or "12pass"
    ' or "1pass2" etc. (or any combination of the same letters)

    For K = 1 To Len(Password)
        SK = SK + (((K Mod 256) Xor Asc(Mid$(Password, K, 1))) Xor Fix(256 * Rnd))
    Next K

    ' init randomizer for encryption/decryption
    Rnd -1
    Randomize SK

    ' encrypt/decrypt every character using the randomizer
    For K = 1 To Len(Str)
        Mid$(Str, K, 1) = Chr(Fix(256 * Rnd) Xor Asc(Mid$(Str, K, 1)))
    Next K

    CryptData = Str
End Function

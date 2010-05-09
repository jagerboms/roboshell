Option Explicit On
Option Strict On

Public Class ScriptOptions
    Private bPKShowName As Boolean = True

    Private bDefShowName As Boolean = True
    Private bDefFix As Boolean = True

    Private bCollShow As Boolean = False

    Private bChkShowName As Boolean = True

    Public Property PrimaryKeyShowName() As Boolean
        Get
            PrimaryKeyShowName = bPKShowName
        End Get
        Set(ByVal pkn As Boolean)
            bPKShowName = pkn
        End Set
    End Property

    Public Property DefaultShowName() As Boolean
        Get
            DefaultShowName = bDefShowName
        End Get
        Set(ByVal dsn As Boolean)
            bDefShowName = dsn
        End Set
    End Property

    Public Property DefaultFix() As Boolean
        Get
            DefaultFix = bDefFix
        End Get
        Set(ByVal df As Boolean)
            bDefFix = df
        End Set
    End Property

    Public Property CollationShow() As Boolean
        Get
            CollationShow = bCollShow
        End Get
        Set(ByVal sc As Boolean)
            bCollShow = sc
        End Set
    End Property

    Public Property CheckShowName() As Boolean
        Get
            CheckShowName = bChkShowName
        End Get
        Set(ByVal dsn As Boolean)
            bChkShowName = dsn
        End Set
    End Property
End Class

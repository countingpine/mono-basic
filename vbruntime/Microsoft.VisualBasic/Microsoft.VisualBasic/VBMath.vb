'
' VBMath.vb
'
' Author:
'   Chris J Breisch (cjbreisch@altavista.net) 
'   Francesco Delfino (pluto@tipic.com)
'   Mizrahi Rafael (rafim@mainsoft.com)
'   Matthew Fearnley (matthew.w.fearnley@gmail.com)
'

'
' Copyright (C) 2002-2006 Mainsoft Corporation.
' Copyright (C) 2004-2006 Novell, Inc (http://www.novell.com)
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the
' "Software"), to deal in the Software without restriction, including
' without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to
' permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
' 
' The above copyright notice and this permission notice shall be
' included in all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
' LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
' OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
' WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'

Imports System
Imports Microsoft.VisualBasic.CompilerServices

Namespace Microsoft.VisualBasic
    <StandardModule()> _
    Public NotInheritable Class VBMath

        ' Declarations
        ' Constructors
        ' Properties
        Private Shared m_seed As Int32
        ' Methods
        Public Shared Function Rnd() As Single
            m_seed = (m_seed * &HFD43FD& + &HC39EC3) And &HFFFFFF
            Return m_seed / &H1000000
        End Function
        Public Shared Function Rnd(ByVal Number As Single) As Single
            If Number = 0.0 Then
                Return m_seed / &H1000000
            ElseIf Number < 0.0 Then
                Dim n As Int32 = BitConverter.ToInt32(BitConverter.GetBytes(Number), 0)
                m_seed = (n And &HFFFFFF) + ((n >> 24) And &HFF)
                Return Rnd()
            End If
            Return Rnd()
        End Function
        Public Shared Sub Randomize()
            Randomize(Timer)
        End Sub
        Public Shared Sub Randomize(ByVal Number As Double)
            Dim n As Int32 = (BitConverter.DoubleToInt64Bits(Number) >> 32)
            n = n Xor n >> 16
            m_seed = (n << 8 And &HFFFF00) Or (m_seed And &HFF)
        End Sub
        ' Events
    End Class
End Namespace

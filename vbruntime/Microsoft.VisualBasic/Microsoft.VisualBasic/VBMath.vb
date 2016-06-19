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

            ' 24-bit linear congruential generation:
            m_seed = (m_seed * &HFD43FD& + &HC39EC3) And &HFFFFFF

            ' Divide m_seed to get into [0,1) range
            Return m_seed / &H1000000

        End Function

        Public Shared Function Rnd(ByVal Number As Single) As Single

            If Number = 0.0 Then
                ' Divide m_seed to get into [0,1) range
                ' Note: m_seed should already be in range [0,2^24)
                Return m_seed / &H1000000

            ElseIf Number < 0.0 Then

                ' Convert Single 'number' bits to Int32 'n':
                Dim n As Int32 = BitConverter.ToInt32(BitConverter.GetBytes(Number), 0)

                ' Add high 8 bits to low 24 bits (similar to MOD 2^24-1, except where hi8+lo24 >= 2^24-1)
                ' Note: m_seed will be cropped to 24 bits in Rnd()
                m_seed = n + ((n >> 24) And &HFF)

            End If

            Return Rnd()

        End Function

        Public Shared Sub Randomize()

            Randomize(Timer)

        End Sub

        Public Shared Sub Randomize(ByVal Number As Double)

            ' Bitwise-convert Double 'Number' to Int32 'n', discarding lower 32 bits (i.e. from the mantissa):
            Dim n As Int32 = (BitConverter.DoubleToInt64Bits(Number) >> 32)

            ' XOR low 16 bits with high 16 bits
            n = n Xor (n >> 16)

            ' Put result into upper 16 bits of the seed, retaining lower 8 bits from previous seed
            m_seed = ((n << 8) And &HFFFF00) Or (m_seed And &HFF)

        End Sub

        ' Events

    End Class
End Namespace

' 
' Visual Basic.Net Compiler
' Copyright (C) 2004 - 2010 Rolf Bjarne Kvinge, RKvinge@novell.com
' 
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
' 
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
' 

Public Enum TokenType
    None

    CharLiteral
    DateLiteral
    DecimalLiteral
    SingleLiteral
    DoubleLiteral
    StringLiteral
    Int16Literal
    Int32Literal
    Int64Literal
    UInt16Literal
    UInt32Literal
    UInt64Literal

    EndOfCode
    EndOfFile
    EndOfLine

    Identifier
    Keyword
    Symbol

End Enum

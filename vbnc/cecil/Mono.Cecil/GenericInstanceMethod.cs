//
// GenericInstanceMethod.cs
//
// Author:
//   Jb Evain (jbevain@gmail.com)
//
// Copyright (c) 2008 - 2010 Jb Evain
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

using System;
using System.Text;

using Mono.Collections.Generic;

namespace Mono.Cecil {

	public sealed class GenericInstanceMethod : MethodSpecification, IGenericInstance, IGenericContext {

		Collection<TypeReference> arguments;
		private Collection<ParameterDefinition> m_resolvedParameters;
		private MethodReturnType m_resolvedReturnType;

		public override MethodReturnType ResolvedReturnType
		{
			get
			{
				if (m_resolvedReturnType == null)
					m_resolvedReturnType = MethodReturnType.ResolveGenericTypes (GenericParameters, GenericArguments);
				return m_resolvedReturnType;
			}
		}

		public override Collection<ParameterDefinition> ResolvedParameters
		{
			get
			{
				if (m_resolvedParameters == null) {
					if (GenericArguments.Count > 0 && GenericParameters.Count == 0) {
						if (OriginalMethod == null) {
							m_resolvedParameters = ParameterDefinitionCollection.ResolveGenericTypes (Parameters, ElementMethod.GenericParameters, GenericArguments);
						} else {
							m_resolvedParameters = ParameterDefinitionCollection.ResolveGenericTypes (Parameters, OriginalMethod.GenericParameters, GenericArguments);
						}
					} else {
						m_resolvedParameters = ParameterDefinitionCollection.ResolveGenericTypes (Parameters, GenericParameters, GenericArguments);
					}
				}
				return m_resolvedParameters;
			}
		}
		public bool HasGenericArguments {
			get { return !arguments.IsNullOrEmpty (); }
		}

		public Collection<TypeReference> GenericArguments {
			get {
				if (arguments == null)
					arguments = new Collection<TypeReference> ();

				return arguments;
			}
		}

		public override bool IsGenericInstance {
			get { return true; }
		}

		IGenericParameterProvider IGenericContext.Method {
			get { return ElementMethod; }
		}

		IGenericParameterProvider IGenericContext.Type {
			get { return ElementMethod.DeclaringType; }
		}

		internal override bool ContainsGenericParameter {
			get { return this.ContainsGenericParameter () || base.ContainsGenericParameter; }
		}

		public override string FullName {
			get {
				var signature = new StringBuilder ();
				var method = this.ElementMethod;
				signature.Append (method.ReturnType.FullName);
				signature.Append (" ");
				signature.Append (method.DeclaringType.FullName);
				signature.Append ("::");
				signature.Append (method.Name);
				this.GenericInstanceFullName (signature);
				this.MethodSignatureFullName (signature);
				return signature.ToString ();

			}
		}

		public GenericInstanceMethod (MethodReference method)
			: base (method)
		{
		}
	}
}

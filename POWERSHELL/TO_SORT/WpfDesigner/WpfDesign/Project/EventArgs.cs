﻿// Copyright (c) 2019 AlphaSierraPapa for the SharpDevelop Team
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

using System;
using System.Collections.Generic;

namespace ICSharpCode.WpfDesign
{
	/// <summary>
	/// Event arguments specifying a component as parameter.
	/// </summary>
	public class DesignItemEventArgs : EventArgs
	{
		readonly DesignItem _item;

		/// <summary>
		/// Creates a new ComponentEventArgs instance.
		/// </summary>
		public DesignItemEventArgs(DesignItem item)
		{
			_item = item;
		}
		
		/// <summary>
		/// The component affected by the event.
		/// </summary>
		public DesignItem Item {
			get { return _item; }
		}
	}
	
	/// <summary>
	/// Event arguments specifying a component and property as parameter.
	/// </summary>
	public class DesignItemPropertyChangedEventArgs : DesignItemEventArgs
	{
		readonly DesignItemProperty _itemProperty;
		readonly object _oldValue;
		readonly object _newValue;

		/// <summary>
		/// Creates a new ComponentEventArgs instance.
		/// </summary>
		public DesignItemPropertyChangedEventArgs(DesignItem item, DesignItemProperty itemProperty) : base(item)
		{
			_itemProperty = itemProperty;
		}

		/// <summary>
		/// Creates a new ComponentEventArgs instance.
		/// </summary>
		public DesignItemPropertyChangedEventArgs(DesignItem item, DesignItemProperty itemProperty, object oldValue, object newValue) : this(item, itemProperty)
		{
			_oldValue = oldValue;
			_newValue = newValue;
		}

		/// <summary>
		/// The property affected by the event.
		/// </summary>
		public DesignItemProperty ItemProperty {
			get { return _itemProperty; }
		}

		/// <summary>
		/// Previous Value
		/// </summary>
		public object OldValue
		{
			get { return _oldValue; }
		}

		/// <summary>
		/// New Value
		/// </summary>
		public object NewValue
		{
			get { return _newValue; }
		}
	}
	
	/// <summary>
	/// Event arguments specifying a component as parameter.
	/// </summary>
	public class DesignItemCollectionEventArgs : EventArgs
	{
		readonly ICollection<DesignItem> _items;

		/// <summary>
		/// Creates a new ComponentCollectionEventArgs instance.
		/// </summary>
		public DesignItemCollectionEventArgs(ICollection<DesignItem> items)
		{
			_items = items;
		}
		
		/// <summary>
		/// The components affected by the event.
		/// </summary>
		public ICollection<DesignItem> Items {
			get { return _items; }
		}
	}
}

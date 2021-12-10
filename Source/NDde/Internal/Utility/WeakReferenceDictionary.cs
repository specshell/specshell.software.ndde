#region Copyright (c) 2005 by Brian Gideon (briangideon@yahoo.com)

/* Shared Source License for NDde
 *
 * This license governs use of the accompanying software ('Software'), and your use of the Software constitutes acceptance of this license.
 *
 * You may use the Software for any commercial or noncommercial purpose, including distributing derivative works.
 *
 * In return, we simply require that you agree:
 *  1. Not to remove any copyright or other notices from the Software.
 *  2. That if you distribute the Software in source code form you do so only under this license (i.e. you must include a complete copy of this
 *     license with your distribution), and if you distribute the Software solely in object form you only do so under a license that complies with
 *     this license.
 *  3. That the Software comes "as is", with no warranties.  None whatsoever.  This means no express, implied or statutory warranty, including
 *     without limitation, warranties of merchantability or fitness for a particular purpose or any warranty of title or non-infringement.  Also,
 *     you must pass this disclaimer on whenever you distribute the Software or derivative works.
 *  4. That no contributor to the Software will be liable for any of those types of damages known as indirect, special, consequential, or incidental
 *     related to the Software or this license, to the maximum extent the law permits, no matter what legal theory it's based on.  Also, you must
 *     pass this limitation of liability on whenever you distribute the Software or derivative works.
 *  5. That if you sue anyone over patents that you think may apply to the Software for a person's use of the Software, your license to the Software
 *     ends automatically.
 *  6. That the patent rights, if any, granted in this license only apply to the Software, not to any derivative works you make.
 *  7. That the Software is subject to U.S. export jurisdiction at the time it is licensed to you, and it may be subject to additional export or
 *     import laws in other places.  You agree to comply with all such laws and regulations that may apply to the Software after delivery of the
 *     software to you.
 *  8. That if you are an agency of the U.S. Government, (i) Software provided pursuant to a solicitation issued on or after December 1, 1995, is
 *     provided with the commercial license rights set forth in this license, and (ii) Software provided pursuant to a solicitation issued prior to
 *     December 1, 1995, is provided with "Restricted Rights" as set forth in FAR, 48 C.F.R. 52.227-14 (June 1987) or DFAR, 48 C.F.R. 252.227-7013
 *     (Oct 1988), as applicable.
 *  9. That your rights under this License end automatically if you breach it in any way.
 * 10. That all rights not expressly granted to you in this license are reserved.
 */

#endregion

using System.Collections;

namespace NDde.Internal.Utility;

internal sealed class WeakReferenceDictionary<TKey, TValue> : IDictionary<TKey, TValue> where TValue : class
{
    private readonly IDictionary<TKey, WeakReference> _Storage = new Dictionary<TKey, WeakReference>();

    public void Add(TKey key, TValue value)
    {
        Purge();
        _Storage.Add(key, new WeakReference(value));
    }

    public bool ContainsKey(TKey key)
    {
        return _Storage.ContainsKey(key);
    }

    public ICollection<TKey> Keys => _Storage.Keys;

    public bool Remove(TKey key)
    {
        return _Storage.Remove(key);
    }

    public bool TryGetValue(TKey key, out TValue value)
    {
        value = null;
        if (!_Storage.ContainsKey(key)) return false;
        value = _Storage[key].Target as TValue;
        return value != null;
    }

    public ICollection<TValue> Values => new MyValueCollection(this);

    public TValue this[TKey key]
    {
        get
        {
            if (!_Storage.ContainsKey(key)) return null;
            if (_Storage[key].Target is TValue value)
                return value;

            return null;
        }
        set
        {
            if (value != null)
            {
                Purge();
                _Storage[key] = new WeakReference(value);
            }
            else
            {
                _Storage.Remove(key);
            }
        }
    }

    public void Add(KeyValuePair<TKey, TValue> item)
    {
        Purge();
        var (key, value) = item;
        _Storage.Add(key, new WeakReference(value));
    }

    public void Clear()
    {
        _Storage.Clear();
    }

    public bool Contains(KeyValuePair<TKey, TValue> item)
    {
        var (key, _) = item;
        if (!_Storage.ContainsKey(key)) return false;
        return _Storage[key].Target is TValue;
    }

    public void CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex)
    {
        var index = 0;
        foreach (var kvp in this)
        {
            array[arrayIndex + index] = kvp;
            index++;
        }
    }

    public int Count => _Storage.Count;

    public bool IsReadOnly => _Storage.IsReadOnly;

    public bool Remove(KeyValuePair<TKey, TValue> item)
    {
        return false;
    }

    public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator()
    {
        foreach (var (key, weakReference) in _Storage)
        {
            if (weakReference.Target is TValue value)
                yield return new KeyValuePair<TKey, TValue>(key, value);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private void Purge()
    {
        var dead = (from kvp in _Storage where !kvp.Value.IsAlive select kvp.Key).ToList();
        foreach (var key in dead)
            _Storage.Remove(key);
    }

    private sealed class MyValueCollection : ICollection<TValue>
    {
        private readonly WeakReferenceDictionary<TKey, TValue> _Parent;

        public MyValueCollection(WeakReferenceDictionary<TKey, TValue> parent)
        {
            _Parent = parent;
        }

        public void Add(TValue item)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void Clear()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public bool Contains(TValue item)
        {
            return this.Any(value => value == item);
        }

        public void CopyTo(TValue[] array, int arrayIndex)
        {
            var index = 0;
            foreach (var value in this)
            {
                array[arrayIndex + index] = value;
                index++;
            }
        }

        public int Count => _Parent._Storage.Values.Count;

        public bool IsReadOnly => true;

        public bool Remove(TValue item)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public IEnumerator<TValue> GetEnumerator()
        {
            foreach (var wr in _Parent._Storage.Values)
            {
                if (wr.Target is TValue value)
                    yield return value;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    } // class
} // class
// namespace
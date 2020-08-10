/////////////////////////////////////////////////////////////////////////////
// Copyright © by W. T. Block, all rights reserved
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include <map>

using namespace std;

// template class to manage map of pointers to any class
template<class KEY, class TYPE>
class CKeyedCollection
{
// public definitions
public: 
	// pair of key and TYPE*
	typedef pair<KEY,TYPE*> PAIR_KEY_PTR;
	// map of key and TYPE*
	typedef map<KEY,TYPE*> MAP_KEY_PTR;
	
// protected data
protected:
	// map of keyed items
	MAP_KEY_PTR m_mapItems;

// methods
public:
	// number of Items
	inline int count()
	{
		return (int)m_mapItems.size(); 
	}

	// clear all Items from the map
	void clear()
	{	// free up the resources in each item
		for ( PAIR_KEY_PTR item : m_mapItems )
		{
			delete item.second;
		}
		// empty the map
		m_mapItems.clear();
	}
	
	// does the key exist in the map?
	bool exists( KEY key )
	{
		TYPE* value = find( key );
		return value != 0;
	}
	
	// find a key in the map
	TYPE* find( KEY key )
	{
		const MAP_KEY_PTR::iterator posEnd = m_mapItems.end();
		const MAP_KEY_PTR::iterator pos = m_mapItems.find( key );
		TYPE* value = 0;
		if ( pos != posEnd )
		{
			value = pos->second;
		}

		return value;
	}
	
	// remove a key from the map
	bool remove( KEY key )
	{
		bool bOK = false;
		TYPE* value = find( key );
		if ( value != 0 )
		{
			delete value;
			m_mapItems.erase( key );
			bOK = true;
		}

		return bOK;
	}
	
	// add a key to the map and return false if it already exists
	bool add( KEY key, TYPE* value )
	{
		const bool bExists = exists( key );
		if ( bExists )
		{
			return false;
		}

		PAIR_KEY_PTR item( key, value );
		m_mapItems.insert( item );
		return true;
	}

// public properties
public:
	// map of keyed items
	inline MAP_KEY_PTR& GetItems()
	{
		return m_mapItems;
	}
	// map of keyed items
	__declspec( property( get=GetItems ))
		MAP_KEY_PTR Items;

	// number of Items
	__declspec( property( get=count ))
		int Count;

	// does the key exist in the map?
	__declspec( property( get=exists ))
		bool Exists[];

// public methods
public:
	// get deleted items returns a map of items that are missing from after
	// that are in before
	static bool GetDeletedItems
	(
		CKeyedCollection<KEY, TYPE>& before, 
		CKeyedCollection<KEY, TYPE>& after,
		CKeyedCollection<KEY, TYPE>& deleted
	)
	{
		bool value = false;
		for ( auto& node : before.Items )
		{
			const KEY key = node.first;
			if ( !after.Exists[ key ] )
			{
				TYPE* temp = new TYPE( *node.second );
				deleted.add( node.first, temp );
			}
		}

		value = deleted.Count > 0;
		return value;
	}

	// get new items returns a map of items that are missing from before
	// that are in after
	static bool GetNewItems
	(
		CKeyedCollection<KEY, TYPE>& before,
		CKeyedCollection<KEY, TYPE>& after,
		CKeyedCollection<KEY, TYPE>& added
	)
	{
		bool value = false;
		for ( auto& node : after.Items )
		{
			const KEY key = node.first;
			if ( !before.Exists[ key ] )
			{
				TYPE* temp = new TYPE( *node.second );
				added.add( node.first, temp );
			}
		}

		value = added.Count > 0;
		return value;
	}

// public construction / destruction
public:
	// constructor
	CKeyedCollection( void )
	{
	}
	// destructor
	virtual ~CKeyedCollection( void )
	{
		clear();
	}
};

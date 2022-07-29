// Handy utility functions for working with spreadsheet data from Google Sheets

function unpivot( source_range, headers_range ) {  // unpivots a data table based on headers_range
    // source_range: first row must be headers
    // headers_range should be a single row range
    // returns a source_range as a taller and thinner table, with 
    //    - a column named variable, filled with variable names from headers_range, and
    //    - a column named value, filled with values corresponding to each variable
    
    var data = range_to_objects( source_range );
    var headers = range_to_array( headers_range )[0];
    var desired_keys = Object.keys( data[0] ).filter( key => !headers.includes( key ) );
  
    let result = data.map( row => {
      let filtered_row = filter_keys( row, desired_keys );
      return headers.map( header => {
        return Object.assign( {}, filtered_row, { variable: header, value: row[header] } );
      } );
    })
    return objects_to_range( result.flat(1) );
  }
  
  function filter_keys( source_object, keys_array ) {  // returns object with subset of properties
    return keys_array.reduce( (obj, key) => { obj[key] = source_object[key]; return obj}, {} )
  }
  
  function range_to_array( range ) {  // converts range to array of arrays
    return range;
  }
  
  function range_to_objects( range ) {  // converts range to array of objects, with properties named after headers
    try {
      var rows = range_to_array( range );
      var keys = rows.splice(0,1)[0];  // side effect: rows loses the first item
      var data = rows.map( row => 
        Object.assign( {}, ...keys.map( (key,i) => ( { [key]: row[i] }) ) ) 
      );
      return data;
    } catch (e) { throw `Range ${range} not structured as an array of objects` }
  }
  
  function objects_to_range( data ) {  // data is a structured array of objects
    if ( data.length < 1 ) throw 'objects_to_range: array of objects is empty';
    headers = Object.keys( data[0] );
    var body = data.map( row => Object.values( row ) )
    new_arr = [headers].concat(body);
    return new_arr;
  }

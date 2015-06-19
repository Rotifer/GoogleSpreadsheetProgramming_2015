// Code Example 6.1
/**
* Return a MySQL connection for given parameters.
*
* This function will create and return a connection
*  or throw an exception if it cannot create a connection
*  using the given parameters.
* 
* @param {String} host
* @param {String} user
* @param {String} pwd
* @param {String} dbName
* @param {number} port
* @return {JdbcConnection}
*/
function getConnection(host, user, pwd, dbName, port) {
  var connectionUrl = ('jdbc:mysql://' + host + ':'
                          + port + '/' + dbName),
      connection;
  try { 
    connection = Jdbc.getConnection(connectionUrl, user, pwd);
    return connection
  } catch (e) {
    throw e; 
  }
}
// Code Example 6.2
/**
* Return a JDBC connection for my
* cloud-hosted MySQL
* 
* You will need to modify the variable
* values for your own MySQL database.
* It calls "getConnection()".
*
* @return {JdbcConnection}
*/
function getConnectionToMyDB() {
  var scriptProperties = 
        PropertiesService.getScriptProperties(),
      pwd = 
        scriptProperties.getProperty('MYSQL_CLOUD_PWD'),
      user = 'sql562209',
      dbName = 'sql562209',
      host = 'sql5.freemysqlhosting.net',
      port = 3306,
      connection;
  connection = getConnection(host, user, pwd, dbName, port);
  return connection;
}
/***
* Test getConnectionToMyDB()
* Prints "JdbcConnection" and closes 
* the Connection object.
*
*/
function test_getConnectionToMyDB() {
  var connection = getConnectionToMyDB();
  Logger.log(connection);
  connection.close();
}

// Code Example 6.3
/**
* Create a JDBC connection and
* execute a CREATE TABLE statement.
*
* return {undefined}
*/
function createTable() {
  var connection,
      ddl,
      stmt;
  ddl = 'CREATE TABLE people('
           + 'first_name VARCHAR(50) NOT NULL,\n'
           + 'surname VARCHAR(50) NOT NULL,\n'
           + 'gender VARCHAR(6),\n'
           + 'height_meter FLOAT(3,2),\n'
           + 'dob DATE)';
  try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    stmt.execute(ddl);
    Logger.log('Table created!');
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }  
}


// Code example 6.4
/**
* Execute a single SQL INSERT statement to
* add the table created by code example 6.3.
*
* @return {undefined}
*/
function populateTable() {
  var connection,
      stmt,
      sql,
      rowInsertedCount;
  sql = "INSERT INTO people(first_name, surname, gender, height_meter, dob) VALUES\n"
            + "('John', 'Browne', 'Male', 1.81, '1980-05-03'),\n"
            + "('Rosa', 'Hernandez', 'Female', 1.70, '1981-04-30'),\n"
            + "('Mary', 'Carr', 'Male', 1.72, '1982-08-01'),\n"
            + "('Lee', 'Chang', 'Male', 1.88, '1973-07-15')";
  try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    rowInsertedCount = stmt.executeUpdate(sql);
    Logger.log(rowInsertedCount + ' values inserted!');
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Code Example 6.5
/**
* Execute a SELECT statement to
* return all table rows
*
* Print rows to log viewer.
* @return {undefined}
*/
function selectFromTable() {
  var connection,
      stmt,
      sql,
      rs,
      rsmd,
      i,
      colCount;
  sql = 'SELECT * FROM people';
  try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    rs = stmt.executeQuery(sql);
    rsmd = rs.getMetaData();
    colCount = rsmd.getColumnCount();
    while(rs.next()) {
      for(i = 1; i <= colCount; i += 1) {
        Logger.log(rs.getString(i) + ', ');
      }
      Logger.log('\n');
    }
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}


// Code Example 6.6
/**
* Execute an UPDATE statement.
*
*
* @return {undefined}
*/
function updateTable() {
  var connection,
      stmt,
      sql,
      updateRowCount;
  sql = "UPDATE people SET gender = 'Female'\n"
       + "WHERE first_name = 'Mary' AND surname = 'Carr'";
    try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    updateRowCount = stmt.executeUpdate(sql);
    Logger.log(updateRowCount + ' rows updated!');
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Code Example 6.7
/**
* Delete a row from the database
*
* Print the deleted row count to
* the log.
*
* @return {undefined}
*/
function deleteTableRow() {
  var connection,
      stmt,
      sql,
      deleteRowCount;
  sql = "DELETE FROM people\n"
       + "WHERE first_name = 'Lee' AND surname = 'Chang'";
    try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    deleteRowCount = stmt.executeUpdate(sql);
    Logger.log('Deleted row count = ' + deleteRowCount);
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Code Example 6.8
/**
* Drop the table from the database
*
* @return {undefined}
*/
function dropTable() {
  var connection,
      stmt,
      ddl;
  ddl = 'DROP TABLE IF EXISTS people';
  try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    stmt.execute(ddl);
    Logger.log('Table "people" dropped!');
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Code Example 6.9
/**
* Demonstrate the use of a prepared statement.
* 
* Prints the name of an emplyee for a 
*  given parameter.
* @return {undefined}
*/
function runPrepStmt() {
  var connection,
      pStmt,
      sql,
      rs,
      empName;
  sql = 'SELECT ename FROM emp WHERE empno = ?';
  try {
    connection = getConnectionToMyDB();
    pStmt = connection.prepareStatement(sql);
    pStmt.setInt(1, 7839);
    rs = pStmt.executeQuery();
    rs.next();
    empName = rs.getString(1);
    Logger.log(empName);
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Code Example 6.10
/**
* Show effect if switching off auto-commit.
*
* Give emplyee number 7900 ("JAMES") a 
* 100-fold pay increase!
* But the change is "rolled back".
* Try changing "connection.rollback()" to
*  "connection.commit()" and see the effect.
*/
function demoTransactions() {
  var connection,
      stmt,
      sql = 'UPDATE emp SET sal = sal * 100 WHERE empno = 7900';
  connection = getConnectionToMyDB();
  connection.setAutoCommit(false);
  stmt = connection.createStatement();
  stmt.executeUpdate(sql);
  connection.rollback();
  connection.close();
}


// Code Example 6.11 - see file "ch06.sql".

// Code Example 6.12
/**
*Print some information about
* the target database to the logger.
*
* @return {undefined}
*/
function printConnectionMetaData() {
  var connection = getConnectionToMyDB(),
      metadata = connection.getMetaData();
  Logger.log('Major Version: ' + metadata.getDatabaseMajorVersion());
  Logger.log('Minor Version: ' + metadata.getDatabaseMinorVersion());
  Logger.log('Product Name: '  + metadata.getDatabaseProductName());
  Logger.log('Product Version: ' + metadata.getDatabaseProductVersion());
  Logger.log('Supports transactions: ' + metadata.supportsTransactions());
}

// Code Example 6.13
/**
* Return a ResultSet
*
* @param {JdbcConnection} conn
* @param {number} deptno
* @return {JdbcResultSet}
*/
function getResSetForDept(conn, deptno) {
  var sql = ('SELECT\n' +
              '  e.ename,\n' +
              '  e.hiredate,\n' +
              '  e.job,\n' +
              '  e.sal,\n' +
              '  d.dname\n' + 
              'FROM\n' +
              '  emp e\n' +
              '  JOIN dept d ON e.deptno = d.deptno\n' +
              'WHERE\n' +
              '  d.deptno = ?'),
      pStmt = conn.prepareStatement(sql),
      rs;
  pStmt.setInt(1, deptno);
  recset = pStmt.executeQuery();
  return recset;
}

// code Example 6.14
/**
* Return the column names for
* given ResultSet
* @param {JdbcResultSet} rs
* @return {String[]}
*/
function getColNames(rs) {
  var md = rs.getMetaData(),
      colCount = md.getColumnCount(),
      colNames = [],
      colName,
      i;
  for(i = 1; i <= colCount; i +=1) {
    colName = md.getColumnName(i);
    colNames.push(colName);
  }
  return colNames;
}
// code Example 6.15
/**
* Return the data rows as an array-of-arrays
* for the given ResultSet
* @param {JdbcResultSet} rs
* @return {String[]}
*/
function getDataRows(rs) {
  var md = rs.getMetaData(),
      colCount = md.getColumnCount(),
      row = [],
      value,
      dataRows = [],
      colName,
      i;
  while(rs.next()) {
    row = [];
    for(i = 1; i <= colCount; i +=1) {
      value = rs.getString(i);
      row.push(value);
    }
    dataRows.push(row);
  }
  return dataRows;  
}
// code Example 6.16
/**
* Check if given sheet name
* exists in the active spreadsheet.
* @param {String} sheetName
* @return {boolean}
*/
function sheetExists(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets(),
      i;
  for(i = 0; i < sheets.length; i += 1) {
    if( sheets[i].getSheetName() == sheetName) {
      return true;
    }
  }
  return false;
}

// Code Example 6.17
/**
* Delete sheet with given name
* and replace it with new one
* with same name.
* @param {String} sheetName
* @return {Sheet}
*/
function replaceOldSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName(sheetName);
  ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);
  return sh
}
// Code Example 6.18
/**
* Create a sheet with the given name
* in the active spreadsheet
* @param {String} sheetName
* @return {Sheet}
*/
function addNewSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.insertSheet(sheetName);
  return sh;
}
// Code Example 6.19
/**
* Run application to extract a set of
* rows from the database and write them
* to a new sheet.
* @return {undefined}
*/
function main() {
  var conn = getConnectionToMyDB(),
      deptno = 20,
      rs = getResSetForDept(conn, deptno),
      colNames = getColNames(rs),
      dataRows = getDataRows(rs),
      sheetName = 'RESEARCH',
      sh,
      dataRowCount = dataRows.length,
      dataColCount = colNames.length,
      outputHeaderRng,
      outputDataRng;
  if(sheetExists(sheetName)) {
     sh = replaceOldSheet(sheetName)
  } else {
    sh = addNewSheet(sheetName);
  }
  outputHeaderRng = sh.getRange(1,1,1,dataColCount);
  outputHeaderRng.setFontWeight('bold');
  outputDataRng = sh.getRange(2, 
                          1, 
                          dataRowCount, 
                          dataColCount);
  outputHeaderRng.setValues([colNames]);
  outputDataRng.setValues(dataRows);
  rs.close();
  conn.close();
}

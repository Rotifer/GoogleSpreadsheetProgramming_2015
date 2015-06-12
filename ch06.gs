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
*
*
*/
function populateTable() {
  var connection,
      stmt,
      sql;
  sql = "INSERT INTO people(first_name, surname, gender, height_meter, dob) VALUES\n"
            + "('John', 'Browne', 'Male', 1.81, '1980-05-03'),\n"
            + "('Rosa', 'Hernandez', 'Female', 1.70, '1981-04-30'),\n"
            + "('Mary', 'Carr', 'Male', 1.72, '1982-08-01'),\n"
            + "('Lee', 'Chang', 'Male', 1.88, '1973-07-15')";
  try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    stmt.execute(sql);
    Logger.log('Values inserted!');
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
      sql;
  sql = "UPDATE people SET gender = 'Female'\n"
       + "WHERE first_name = 'Mary' AND surname = 'Carr'";
    try {
    connection = getConnectionToMyDB();
    stmt = connection.createStatement();
    // Could also use executeUpdate() method here.
    stmt.execute(sql);
    Logger.log('Table updated!');
  } catch (e) {
    Logger.log(e);
    throw e;
  } finally {
    connection.close();
  }
}

// Delete row
// Exception: Can not issue data manipulation statements with executeQuery().
// First run gives log: Deleted row count = 1
// Second run gives log:  Deleted row count = 0
// http://code.rohitink.com/2013/06/12/mysql-integer-float-decimal-data-types-differences/

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
    // Using executeUpdate() method here.
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


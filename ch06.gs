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
  var scriptProperties = PropertiesService.getScriptProperties(),
      pwd = scriptProperties.getProperty('MYSQL_CLOUD_PWD'),
      user = 'sql562209',
      dbName = 'sql562209',
      host = 'sql5.freemysqlhosting.net',
      port = 3306;  
}

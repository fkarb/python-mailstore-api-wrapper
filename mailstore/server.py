# -*- coding: utf-8 -*-
#
# Copyright (c) 2012, 2013, 2014 MailStore Software GmbH
#
# Permission is hereby granted, free of charge, to any person obtaining
# a copy of this software and associated documentation files (the
# "Software"), to deal in the Software without restriction, including
# without limitation the rights to use, copy, modify, merge, publish,
# distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so, subject to
# the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

__doc__ = """Wrapper for MailStore Server's Administration API"""

import urllib.request
import urllib.error
import urllib.parse
import json
import mailstore.errors

class Client():
    """The API client class"""
    def __init__(self,
                 username = "admin",
                 password = "admin",
                 host = "127.0.0.1",
                 port = 8463,
                 autoHandleToken = True,
                 waitTime = 1000,
                 callbackStatus = None,
                 logLevel = 2):

        # Initialize connection settings
        self.username = username
        self.password = password
        self.host = host
        self.port = port
  
        # If set to true, client handles tokens/long running tasks itself.
        self.autoHandleToken = autoHandleToken 

        # Time in milliseconds the API should wait before returning a status token.
        self.waitTime = waitTime               

        # Define logging parameters
        self.logLevel = logLevel
        self.logLevels = {0: "NONE",     # No log output
                          1: "ERROR",    # Log errors only 
                          2: "WARNING",  # Log errors and warnings
                          3: "INFO",     # Log informational about what is being done
                          4: "DEBUG"}    # Log also send and received data

        # Callback Function for status
        self.callbackStatus = callbackStatus

        # Initialize password manager
        self.passwordMgr = urllib.request.HTTPPasswordMgrWithDefaultRealm()
        self.realm = (None,
                      "https://{}:{}".format(self.host, self.port),
                      self.username,
                      self.password)
        self.passwordMgr.add_password(*self.realm)
        self.authMgr = urllib.request.HTTPBasicAuthHandler(password_mgr=self.passwordMgr)
        self.opener = urllib.request.build_opener(self.authMgr)
        self.installOpener = urllib.request.install_opener(self.opener)

    # ---------------------------------------------------------------- #
    # Private Methods                                                  #
    # ---------------------------------------------------------------- #
    
    def __logprint(self, logLevel, *args):
        """Helper method for printing additional information based on the logLevel. '
        Levels above 2 are especially useful for debug purpose
        """
        
        if logLevel <= self.logLevel:
            print(self.logLevels[logLevel], args)


    def __hasToken(self, jsonValues):
        """Helper method to verify if all required attributes for token handling are available."""
        if "token" in jsonValues and jsonValues["token"] is not None and "statusVersion" in jsonValues:
            self.__logprint(3, "__hasToken: Status token " + jsonValues["token"] + " detected. statusVersion is " + str(jsonValues["statusVersion"]))
            return True
        else:
            self.__logprint(3, "__hasToken: No status token detected")
            return False


    def __handleToken(self, jsonValues, waitTime=None):
        """Helper function for status tokens handling"""

        waitTime = waitTime if waitTime is not None else self.waitTime

        # Execute callback function for initial state
        if callable(self.callbackStatus):
            self.__logprint(3, "__handleToken: Executing callback function \"" + self.callbackStatus.__name__ + "\" for first status.")
            self.callbackStatus(jsonValues)
 
        while jsonValues["statusCode"] == "running":
            self.__logprint(3, "__handleToken: Refreshing status for task with token " + jsonValues["token"] + ".")
            jsonValues = self.GetStatus(jsonValues, waitTime=waitTime)
            self.__logprint(4, "__handleToken:", jsonValues)

            # Execute callback function for subsequent and final state
            if callable(self.callbackStatus):
                self.__logprint(3, "__handleToken: Executing callback function \"" + self.callbackStatus.__name__ + "\" for refreshed status.")
                self.callbackStatus(jsonValues)

        self.__logprint(3, "__handleToken: Task with token " + jsonValues["token"] + " finished.")
        return jsonValues


    def __callMethod(self, method, arguments = {}, mode = "invoke", autoHandleToken = None):
        """This is where the magic happens! Method is called by all other public methods that wrap 
        an Administration API method."""

        autoHandleToken = autoHandleToken if autoHandleToken is not None else self.autoHandleToken

        url = "https://{}:{}/api/{}/{}".format(self.host, self.port, mode, method)
        data = urllib.parse.urlencode([(key, arguments[key]) for key in list(arguments) if arguments[key]])

        self.__logprint(4, "__callMethod: METHOD:", method)
        self.__logprint(4, "__callMethod: ARGUMENTS:", arguments)
        self.__logprint(3, "__callMethod: HTTP POST:", url, data)

        # Try making the HTTP request...
        try:
            response = urllib.request.urlopen(url, data=data.encode())
        # ...and catch exceptions.
        except urllib.error.HTTPError as e:
            exceptionString = "{} {} {} {} {}".format(e.code, e.msg, url, e.fp._method, data)
            self.__logprint(1, exceptionString)
            raise e
        except Exception as e:
            self.__logprint(1, "Unhandled Exception")
            raise mailstore.errors.MailStoreBaseError(e)

        # Parse server response, which is always in JSON format.
        decodedValues = response.read().decode("utf-8-sig")
        jsonValues = json.loads(decodedValues)
        self.__logprint(4, "__callMethod: HTTP RESPONSE:", decodedValues)

        # Check if response contains a status token and, depending on the
        # value of autoHandleToken, handle the token ourselves or just
        # return the JSON response to the caller.
        if self.__hasToken(jsonValues):
            if autoHandleToken:
                self.__logprint(3, "__callMethod: Automatic token handling is ENABLED.")
                returnData = self.__handleToken(jsonValues)
            else:
                self.__logprint(3, "__callMethod: Automatic token handling is DISABLED.")
                returnData = jsonValues
        else:
            returnData = jsonValues

        self.__logprint(3, "__callMethod: Returning data to caller \"" + method + "\"")
        self.__logprint(4, "__callMethod: ", returnData)

        return returnData


    # ---------------------------------------------------------------- #
    # Public Methods                                                   #
    # ---------------------------------------------------------------- #
 
    def GetStatus(self, jsonValues, waitTime=None):
        """Retrieve and update status token of long running task. This
        method is used for automatic token handling, but can also be
        called directly when manual token handling is done."""
        
        waitTime = waitTime if waitTime else self.waitTime

        if self.__hasToken(jsonValues):
            statusVersion = str(jsonValues["statusVersion"])
            jsonValues = self.__callMethod("get-status", {"token": jsonValues["token"], "millisecondsTimeout": waitTime, "lastKnownStatusVersion": statusVersion}, mode="", autoHandleToken=False)
            return jsonValues
        else:
            self.__logprint(1, "GetStatus: Cannot get status, no token found!")
            raise mailstore.errors.MailStoreNoTokenError(jsonValues)

    def CancelAsync(self, jsonValues):
        """Cancels a long running task."""
        if self.__hasToken(jsonValues):
            return self.__callMethod("cancel-async", {"token": jsonValues["token"]}, mode="")
        else:
            self.__logprint(1, "CancelAsync: Cannot cancel, no token found!")
            raise mailstore.errors.MailStoreNoTokenError(jsonValues)


    # ---------------------------------------------------------------- #
    # Wrapped Administration API methods                               #
    # ---------------------------------------------------------------- #
 
    def AttachStore(self, name, type=None, databasePath=None, contentPath=None, indexPath=None,
                    serverName=None, userName=None, password=None, databaseName=None, requestedState=None, autoHandleToken=None):
        """Attaches an existing archive store
        
        name:            A meaningful name for the archive store. Examples: "Messages 2012" or "2012-01".
        type:            Type of archive store. Must be one of the following:
                           * FileSystemStandard
                           * FileSystemAdvanced
                           * FileSystemFlat
                           * SQLServer
                           * PostgreSQL
        databasePath:    Directory containing folder information and email meta data. (FileSystemStandard, FileSystemAdvanced > FileSystemFlat only)
        contentPath:     Directory containing email headers and contents.
        indexPath:       Directory containing the full-text indexes.
        serverName:      Hostname or IP address of database server (MS SQL Server and PostgreSQL only)
        userName:        Username for database access (MS SQL Server and PostgreSQL only)
        password:        Password for database access MS SQL Server and PostgreSQL only)
        databaseName:    Name of SQL database containing folder information and e-mail metadata.
        requestedState:  Status of the archive store after attaching. Must be one of the follwing 
                           * current         New email messages should be archived into this store.
                           * normal          The archive store should be opened normally. Write access is possible, but new email messages are not archived into this store.
                           * writeProtected  The archive store should be write-protected.
                           * disabled        The archive store should be disabled. This causes the archive store to be closed if it is currently open.
        """
        return self.__callMethod("AttachStore", {"name": name, "type": type, "databaseName": databaseName, "databasePath": databasePath,
                                         "contentPath": contentPath, "indexPath": indexPath, "serverName": serverName,
                                         "userName": userName, "password": password, "requestedState": requestedState}, autoHandleToken=autoHandleToken)


    def ClearUserPrivilegesOnFolders(self, userName, autoHandleToken=None):
        """ Removes all privileges that a user has on archive folders.

        userName:  The user name of the user whose privileges on archive folders should be removed."""
        return self.__callMethod("ClearUserPrivilegesOnFolders", {"userName": userName}, autoHandleToken=autoHandleToken)


    def CompactMasterDatabase(self, autoHandleToken=None):
        """Compacts the master database"""
        return self.__callMethod("CompactMasterDatabase", autoHandleToken=autoHandleToken)


    def CompactStore(self, id, autoHandleToken=None):
        """Compacts an archive store

        id:  The uniqe identifier of the archive store to be compacted."""
        return self.__callMethod("CompactStore", {"id": id}, autoHandleToken=autoHandleToken)


    def CreateProfile(self, properties=None, raw=True, autoHandleToken=None):
        """Create a new archiving or exporting profile

        properties:  The raw profile properties. Values of an existing profile can be used as template."""
        raw = "true" if raw else "false"
        return self.__callMethod("CreateProfile", {"properties": properties, "raw": raw}, autoHandleToken=autoHandleToken)


    def CreateStore(self, name=None, type=None, databasePath=None, contentPath=None, indexPath=None,
                    serverName=None, userName=None, password=None, databaseName=None, requestedState=None, autoHandleToken=None):
        """Creates a new archive store and attaches it afterwards

        name:            A meaningful name for the archive store. Examples: "Messages 2012" or "2012-01".
        type:            Type of archive store. Must be one of the following:
                           * FileSystemStandard
                           * FileSystemAdvanced
                           * FileSystemFlat
                           * SQLServer
                           * PostgreSQL
        databasePath:    Directory containing folder information and email meta data. (FileSystemStandard, FileSystemAdvanced > FileSystemFlat only)
        contentPath:     Directory containing email headers and contents.
        indexPath:       Directory containing the full-text indexes.
        serverName:      Hostname or IP address of database server (MS SQL Server and PostgreSQL only)
        userName:        Username for database access (MS SQL Server and PostgreSQL only)
        password:        Password for database access MS SQL Server and PostgreSQL only)
        databaseName:    Name of SQL database containing folder information and e-mail metadata.
        requestedState:  Status of the archive store after attaching. Must be one of the follwing 
                           * current         New email messages should be archived into this store.
                           * normal          The archive store should be opened normally. Write access is possible, but new email messages are not archived into this store.
                           * writeProtected  The archive store should be write-protected.
                           * disabled        The archive store should be disabled. This causes the archive store to be closed if it is currently open.
        """
        return self.__callMethod("CreateStore", {"name": name, "type": type, "databasePath": databasePath, "contentPath": contentPath,
                                         "indexPath": indexPath, "serverName": serverName, "userName": userName, "password": password,
                                         "databaseName": databaseName, "requestedState": requestedState}, autoHandleToken=autoHandleToken)


    def CreateUser(self, userName, privileges, fullName=None, distinguishedName=None, authentication=None, password=None, autoHandleToken=None):
        """Create a new user

        userName:           Name of the user to be created.
        privileges:         Comma-separated list of global privileges that the user should be granted. Possible values are:
                              * none                   The user is granted no global privileges.
                                                       If specified, this value has to be the only value in the list.
                              * admin                  The user is granted administrator privileges. 
                                                       If specified, this value has to be the only value in the list.
                              * login                  The user can log on to MailStore Server.
                              * changePassword         The user can change his own MailStore Server password. 
                                                       Only useful if the authentication is set to 'integrated'.
                              * archive                The user can run archiving profiles.
                              * modifyArchiveProfiles  The user can create, modify and delete archiving profiles.
                              * export                 The user can run export profiles.
                              * modifyExportProfiles   The user can create, modify and delete export profiles.
                              * delete:                The user can delete messages. 
                                                       Please note: Normal user can only delete messages in folders where he has 
                                                       been granted delete access. In addition, compliance settings may be in 
                                                       effect, preventing administrators and normal users from deleting messages 
                                                       even when they have been granted the privilege to do so.
        fullName:           (optional) The full name (display name) of the user, e.g. "John Doe".
        distinguishedName:  (optional) The LDAP distinguished name of the user. This is typically automatically
                            specified when synchronizing with Active Directory or other LDAP servers.
        authentication:     (optional) The authentication mode. Possible values are:
        integrated:         Specifies MailStore-integrated authentication. This is the default value.
        directoryServices:  Specified Directory Services authentication. If this value is specified,
                            the password is stored, but is ignored when the user logs on to MailStore Server.
        password:           (optional) The password that the user can use to log on to MailStore Server.
                            Only used when authentication is set 'to integrated'.
        """
        return self.__callMethod("CreateUser", {"userName": userName, "privileges": privileges, "fullName": fullName,
                                        "distinguishedName": distinguishedName, "authentication": authentication,
                                        "password": password}, autoHandleToken=autoHandleToken)


    def DeleteEmptyFolders(self, folder=None, autoHandleToken=None):
        """Deletes archive folders which don't contain any messages

        folder:  (optional) If specified, only this folder and its subfolders are deleted if empty.
                            Folder delimiter is /
        """
        return self.__callMethod("DeleteEmptyFolders", {"folder": folder}, autoHandleToken=autoHandleToken)


    def DeleteMessage(self, id, autoHandleToken=None):
        """Deletes a single message from the archive

        id:  The uniqe identifier of the message to be deleted in format: <store_id>:<message_num>"""
        return self.__callMethod("DeleteMessage", {"id": id}, autoHandleToken=autoHandleToken)


    def DeleteProfile(self, id, autoHandleToken=None):
        """Deletes an archiving or export profile

        id:  The unique identifier of the profile to be deleted."""
        return self.__callMethod("DeleteProfile", {"id": id}, autoHandleToken=autoHandleToken)


    def DeleteUser(self, userName, autoHandleToken=None):
        """Delete a user 

        Neither the user's archive nor the user's archived e-mail are deleted when deleting users.

        userName:  The user name of the user to be deleted.
        """
        return self.__callMethod("DeleteUser", {"userName": userName}, autoHandleToken=autoHandleToken)


    def DetachStore(self, id, autoHandleToken=None):
        """Detache archive store

        id:  This unique identifier of the archive store to be detached.
        """
        return self.__callMethod("DetachStore", {"id": id}, autoHandleToken=autoHandleToken)


    def GetActiveSessions(self, autoHandleToken=None):
        """Retrieve list of active logon sessions"""
        return self.__callMethod("GetActiveSessions", autoHandleToken=autoHandleToken)


    def GetChildFolders(self, folder=None, maxLevels=None, autoHandleToken=None):
        """Retrieves a list of child folders of a specific folder

        folder:     (optional) The folder of which the child folders are to be retrieved. If you don't specify this parameter,
                    the method returns the child folders of the root level (user archives).
        maxLevels:  (optional) If maxLevels is not specified, this method returns the child folders recursively,
                    which means that you get the whole folder hierarchy starting at the folder specified.
                    Set maxLevels to a value equal to or greater than 1 to limit the levels returned.
        """
        return self.__callMethod("GetChildFolders", {"folder": folder, "maxLevels": maxLevels}, autoHandleToken=autoHandleToken)


    def GetComplianceConfiguration(self, autoHandleToken=None):
        """Retrieve the current compliance configuration"""
        return self.__callMethod("GetComplianceConfiguration", autoHandleToken=autoHandleToken)


    def GetDirectoryServicesConfiguration(self, autoHandleToken=None):
        """Retrieve the current directory service configuration"""
        return self.__callMethod("GetDirectoryServicesConfiguration", autoHandleToken=autoHandleToken)


    def GetFolderStatistics(self, autoHandleToken=None):
        """Retrieve folder statistics"""
        return self.__callMethod("GetFolderStatistics", autoHandleToken=autoHandleToken)


    def GetMessages(self, folder, autoHandleToken=None):
        """Retrieve list of messages from a specific folder

        folder:  The folder from which to retrieve the message list
        """
        return self.__callMethod("GetMessages", {"folder" : folder}, autoHandleToken=autoHandleToken)

    
    def GetProfiles(self, raw=True, autoHandleToken=None):
        """Retrieve list of profiles"""
        return self.__callMethod("GetProfiles", {"raw": raw}, autoHandleToken=autoHandleToken)


    def GetServerInfo(self, autoHandleToken=None):
        """Retrieve list of general server information"""
        return self.__callMethod("GetServerInfo", autoHandleToken=autoHandleToken)


    def GetStoreIndexes(self, id, autoHandleToken=None):
        """Retrieve list of full-text indexes for given arechive store

        id:  The unique identifier of the archive store whose full-text indexes are to be returned.
        """
        return self.__callMethod("GetStoreIndexes", {"id": id}, autoHandleToken=autoHandleToken)


    def GetStores(self, autoHandleToken=None):
        """Retrieve a list of attached archive stores"""
        return self.__callMethod("GetStores", autoHandleToken=autoHandleToken)


    def GetTimeZones(self, autoHandleToken=None):
        """Retrieve list of all available time zones on the server 
 
        This is particularly useful for GetWorkerResults method.
        """
        return self.__callMethod("GetTimeZones", autoHandleToken=autoHandleToken)


    def GetUserInfo(self, userName, autoHandleToken=None):
        """Retrieve detailed user information about specific user
 
        userName:  User name of the user whose information should be returned."""
        return self.__callMethod("GetUserInfo",{"userName":userName}, autoHandleToken=autoHandleToken)


    def GetUsers(self, autoHandleToken=None):
        """Retrieve list of all users"""
        return self.__callMethod("GetUsers", autoHandleToken=autoHandleToken)


    def GetWorkerResults(self, fromIncluding, toExcluding, timeZoneID="$Local", profileID=None, userName=None, autoHandleToken=None):
        """Retrieves list of finished profile executions
 
        fromIncluding:  The date which indicates the beginning time, e.g. "2013-01-01T00:00:00".
        toExcluding:    The date which indicates the ending time, e.g. "2013-02-28T23:59:59".
        timeZoneID:     The time zone the date should be converted to, e.g. "$Local",
                        which represents the time zone of the operating system.
        profileID:      The profile id for which to retrieve results.
        userName:       The user name for which to retrieve results.
        """
        return self.__callMethod("GetWorkerResults", {"fromIncluding": fromIncluding, "toExcluding": toExcluding, "timeZoneID": timeZoneID, "profileID": profileID, "userName": userName}, autoHandleToken=autoHandleToken)


    def MaintainFileSystemDatabases(self, autoHandleToken=None):
        """Runs maintenance on all file system-based archive store databases 

        Each Firebird embedded database file will be rebuild by this operation 
        by creating a backup file and restoring from that backup file.
        """
        return self.__callMethod("MaintainFileSystemDatabases", autoHandleToken=autoHandleToken)


    def MergeStore(self, id, sourceId, autoHandleToken=None):
        """Merge two archive stores.

        The source archive store remains unchanged and must be detached afterwards.
        
        id:        Unique identifier of destination archive store
        sourceId:  Unique identifier of source archive store
        """
        return self.__callMethod("MergeStore", {"id" : id, "sourceId" : sourceId}, autoHandleToken=autoHandleToken)


    def MoveFolder(self, fromFolder, toFolder, autoHandleToken=None):
        """Move or rename an archive folder

        fromFolder: The folder which should be moved or renamed, e.g. "johndoe/Outlook/Inbox".
        toFolder:   The target folder name, e.g. "johndoe/Outlook/Inbox-new".

        Examples:
        The following example renames the user archive "johndoe" to "john.doe".

          MoveFolder --fromFolder="johndoe" --toFolder="john.doe"


        The following example renames the folder "Outlook" within the user archive "johndoe" to "Microsoft Outlook".

          MoveFolder --fromFolder="johndoe/Outlook" --toFolder="johndoe/Microsoft Outlook"


        The following example moves the folder "Project A" into the folder "Projects".

          MoveFolder --fromFolder="johndoe/Outlook/Project A" --toFolder="johndoe/Outlook/Projects/Project A
        """
        return self.__callMethod("MoveFolder", {"fromFolder": fromFolder, "toFolder": toFolder}, autoHandleToken=autoHandleToken)


    def RebuildStoreIndex(self, id, folder, autoHandleToken=None):
        """Rebuild full-text index

        id:      The unique identifier of the archive store that contains the full-text index to be rebuilt.
        folder:  Name of the archive of which the full-text index should be rebuild e.g. "johndoe".
        """
        return self.__callMethod("RebuildStoreIndex", {"id": id, "folder": folder}, autoHandleToken=autoHandleToken)


    def RefreshAllStoreStatistics(self, autoHandleToken=None):
        """Refresh statistics of all attached archive stores"""
        return self.__callMethod("RefreshAllStoreStatistics", autoHandleToken=autoHandleToken)


    def RenameStore(self, id, name, autoHandleToken=None):
        """Rename archive store

        id:    The unique identifier of the archive store to be renamed.
        name:  The new archive store name.
        """
        return self.__callMethod("RenameStore", {"id": id, "name": name}, autoHandleToken=autoHandleToken)


    def RenameUser(self, oldUserName, newUserName, autoHandleToken=None):
        """Rename user. 

        Note tht the user's archive will not be renamed by this method.

        oldUserName:  User name of the user to be renamed.
        newUserName:  New user name.
        """
        return self.__callMethod("RenameUser", {"oldUserName": oldUserName, "newUserName": newUserName}, autoHandleToken=autoHandleToken)


    def RetryOpenStores(self, autoHandleToken=None):
        """Retry opening stores that could not be opened the last time"""
        return self.__callMethod("RetryOpenStores", autoHandleToken=autoHandleToken)

    def RunTemporaryProfile(self, properties=None, raw=True, autoHandleToken=None):
        """Run temporary archiving or exporting profile

        Using this method will run the profile once, without actually storing the profile configuration in the database.

        properties:  The raw profile properties. Values of an existing profile can be used as template
        """
        raw = "true" if raw else "false"
        return self.__callMethod("RunTemporaryProfile", {"properties": properties, "raw": raw}, autoHandleToken=autoHandleToken)

    def RunProfile(self, id, autoHandleToken=None):
        """Run existing archiving or exporting profile

        id:  The identifier of the profile to be run.
        """
        return self.__callMethod("RunProfile", {"id" : id}, autoHandleToken=autoHandleToken)

    def SetComplianceConfiguration(self, config, autoHandleToken=None):
        """Set compliance configuration

        config:  Raw configuration object. Use GetComplianceConfiguration to retrieve a valid object.
        """
        return self.__callMethod("SetComplianceConfiguration", {"config": config}, autoHandleToken=autoHandleToken)

    def SetDirectoryServicesConfiguration(self, config, autoHandleToken=None):
        """Set directory service configuration

        config:  Raw configuration object. Use GetDirectoryServicesConfiguraion to retrieve a valid object.
        """ 
        return self.__callMethod("SetDirectoryServicesConfiguration", {"config" : config}, autoHandleToken=autoHandleToken)

    def SetStoreProperties(self, id, type=None, databasePath=None, contentPath=None, indexPath=None,
                           serverName=None, userName=None, password=None, databaseName=None, autoHandleToken=None):
        """Set properties of a store

        id:              Unique identifier of archive store to be modified.
        type:            Type of archive store. Must be one of the following:
                           * FileSystemStandard
                           * FileSystemAdvanced
                           * FileSystemFlat
                           * SQLServer
                           * PostgreSQL
        databasePath:    Directory containing folder information and email meta data. (FileSystemStandard, FileSystemAdvanced > FileSystemFlat only)
        contentPath:     Directory containing email headers and contents.
        indexPath:       Directory containing the full-text indexes.
        serverName:      Hostname or IP address of database server (MS SQL Server and PostgreSQL only)
        userName:        Username for database access (MS SQL Server and PostgreSQL only)
        password:        Password for database access MS SQL Server and PostgreSQL only)
        databaseName:    Name of SQL database containing folder information and e-mail metadata.
        """
        return self.__callMethod("SetStoreProperties", {"id": id, "type": type, "databaseName": databaseName, "databasePath": databasePath,
                                         "contentPath": contentPath, "indexPath": indexPath, "serverName": serverName,
                                         "userName": userName, "password": password}, autoHandleToken=autoHandleToken)


    def SetStoreRequestedState(self, id, requestedState, autoHandleToken=None):
        """Set requested state of an archive store

        id:              Unique identifier of the archive store whose requested state should be set.
        requestedState:  Status of the archive store after attaching. Must be one of the follwing 
                           * current         New email messages should be archived into this store.
                           * normal          The archive store should be opened normally. Write access is possible, but new email messages are not archived into this store.
                           * writeProtected  The archive store should be write-protected.
                           * disabled        The archive store should be disabled. This causes the archive store to be closed if it is currently open.
        """
        return self.__callMethod("SetStoreRequestedState", {"id": id, "requestedState": requestedState}, autoHandleToken=autoHandleToken)


    def SetUserAuthentication(self, userName, authentication, autoHandleToken=None):
        """Set authentication mode of a user

        userName:        The user name of the user whose authentication mode should be set.
        authentication:  The authentication mode. Possible values are:
                           * integrated          Specifies MailStore-integrated authentication. This is the default value.
                           * directoryServices   Specified Directory Services authentication. If this value is specified,
                                                 the password is stored, but is ignored when the user logs on to MailStore Server.
        """
        return self.__callMethod("SetUserAuthentication", {"userName": userName, "authentication": authentication}, autoHandleToken=autoHandleToken)


    def SetUserDistinguishedName(self, userName, distinguishedName=None, autoHandleToken=None):
        """Set distinguished name (DN) of a user

        userName            The user name of the user whose distinguished name should be set (or removed).
        distinguishedName:  (optional) The distinguished name to be set. If this argument is not specified,
                            the distinguished name of the specified user is removed.
        """
        return self.__callMethod("SetUserDistinguishedName", {"userName": userName, "distinguishedName": distinguishedName}, autoHandleToken=autoHandleToken)


    def SetUserEmailAddresses(self, userName, emailAddresses=None, autoHandleToken=None):
        """Sets the e-mail addresses of a user

        userName:        The user name of the user whose e-mail addresses are to be set.
        emailAddresses:  (optional) A comma-separated list of e-mail addresses. The first e-mail address 
                         in the list must be the user's primary e-mail address.
        """
        if isinstance(emailAddresses, (list, tuple)):
            emailAddresses = ",".join(emailAddresses)
        return self.__callMethod("SetUserEmailAddresses", {"userName": userName, "emailAddresses": emailAddresses}, autoHandleToken=autoHandleToken)


    def SetUserFullName(self, userName, fullName=None, autoHandleToken=None):
        """Set the full name (display name) of a user

        userName:  The user name of the user whose full name (display name) should be set (or removed).
        fullName:  (optional) The full name to be set. If this argument is not specified, the full 
                   name of the specified user is removed.
        """
        return self.__callMethod("SetUserFullName", {"userName": userName, "fullName": fullName}, autoHandleToken=autoHandleToken)


    def SetUserPassword(self, userName, password, autoHandleToken=None):
        """Set password of a user 

        userName:  The user name of the user whose MailStore Server should be set.
        password:  The new password.
        """
        return self.__callMethod("SetUserPassword", {"userName": userName, "password": password}, autoHandleToken=autoHandleToken)


    def SetUserPop3UserNames(self, userName, pop3UserNames=None, autoHandleToken=None):
        """Sets POP3 user names of a user (used for MailStore Proxy).

        userName:       The user name of the user whose POP3 user names should be set.
        pop3UserNames:  (optional) A comma-separated list of POP3 user names that should be set.
        """
        if isinstance(pop3UserNames, (list, tuple)):
            pop3UserNames = ",".join(pop3UserNames)
        return self.__callMethod("SetUserPop3UserNames", {"userName": userName, "pop3UserNames": pop3UserNames}, autoHandleToken=autoHandleToken)


    def SetUserPrivileges(self, userName, privileges, autoHandleToken=None):
        """Set the privileges of a user

        userName:  The user name of the user whose global privileges should be set.
        privileges:         Comma-separated list of global privileges that the user should be granted. Possible values are:
                              * none                   The user is granted no global privileges.
                                                       If specified, this value has to be the only value in the list.
                              * admin                  The user is granted administrator privileges. 
                                                       If specified, this value has to be the only value in the list.
                              * login                  The user can log on to MailStore Server.
                              * changePassword         The user can change his own MailStore Server password. 
                                                       Only useful if the authentication is set to 'integrated'.
                              * archive                The user can run archiving profiles.
                              * modifyArchiveProfiles  The user can create, modify and delete archiving profiles.
                              * export                 The user can run export profiles.
                              * modifyExportProfiles   The user can create, modify and delete export profiles.
                              * delete:                The user can delete messages. 
                                                       Please note: Normal user can only delete messages in folders where he has 
                                                       been granted delete access. In addition, compliance settings may be in 
                                                       effect, preventing administrators and normal users from deleting messages 
                                                       even when they have been granted the privilege to do so.
        """
        if isinstance(privileges, (list, tuple)):
            privileges = ",".join(privileges)
        return self.__callMethod("SetUserPrivileges", {"userName": userName, "privileges": privileges}, autoHandleToken=autoHandleToken)


    def SetUserPrivilegesOnFolder(self, userName, folder, privileges, autoHandleToken=None):
        """Set user's privileges on a specific folder

        userName:    The user name of the user who should be granted or denied privileges.
        folder:      The folder on which the user should be granted or denied privileges.
                     In the current version, this can only be a top-level folder (user archive).
        privileges:  A comma-separated list of privileges that the specified user should be granted on the specified folder. Possible values are:
                       * none    The user is denied access to the specified folder. If specified, this value has to be the only value in the list.
                       * read    The user is granted read access to the specified folder.
                       * write   The user is granted write access to the specified folder.
                       * delete  The user is granted delete access to the specified folder.
        """
        if isinstance(privileges, (list, tuple)):
            privileges = ",".join(privileges)
        return self.__callMethod("SetUserPrivilegesOnFolder", {"userName": userName, "folder": folder, "privileges": privileges}, autoHandleToken=autoHandleToken)


    def SyncUsersWithDirectoryServices(self, dryRun=None, autoHandleToken=None):
        """Synchronizes with currently configured directory service

        dryRun: if set, only retrieve changes from the directory service syncronization
                but do not store them in the user database.
        """
        dryRun = "true" if dryRun else "false"
        return self.__callMethod("SyncUsersWithDirectoryServices", {"dryRun": dryRun}, autoHandleToken=autoHandleToken)


    def UpgradeStore(self, id, autoHandleToken=None):
        """Upgrade archive store 
      
        Only useful for archive stores that have been created in MailStore Server 5.x or earlier.

        id:  The unique identifier of the archive store to be upgraded.
        """
        return self.__callMethod("UpgradeStore", {"id": id}, autoHandleToken=autoHandleToken)


    def VerifyStore(self, id, autoHandleToken=None):
        """Verify content of an archive store.

        id: The uniqe identifier of the archive store to be verified.
        """
        return self.__callMethod("VerifyStore", {"id": id}, autoHandleToken=autoHandleToken)

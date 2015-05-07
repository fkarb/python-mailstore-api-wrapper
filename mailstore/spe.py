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

__doc__ = """Wrapper for MailStore Service Provider Editions's Management API"""

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
                 port = 8474,
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
 
    def AttachStore(self, instanceID, name, path, requestedState=None, autoHandleToken=None):
        """Attach existing archive store.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param name:            Meaningful name of archive store.
        :type name:             str
        :param path:            Path of directory containing archive store data.
        :type path              str
        :param requestedState:  State of archive store after attaching.
        """
        return self.__callMethod("AttachStore", {"instanceID": instanceID, "name": name, "path": path, "requestedState": requestedState}, autoHandleToken=autoHandleToken)

    def ClearUserPrivilegesOnFolders(self, instanceID, userName, autoHandleToken=None):
        """Removes all privileges of a user on all archive folders.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param userName:    User name of MailStore user.
        :type userName:     str
        """
        return self.__callMethod("ClearUserPrivilegesOnFolders", {"instanceID": instanceID, "userName": userName}, autoHandleToken=autoHandleToken)

    def CompactStore(self, instanceID, id, autoHandleToken=None):
        """Compact archive store

        :param instanceID:  Unique ID of MailStore instance in which this command should be invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store
        :type id:           int
        """
        return self.__callMethod("CompactStore", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def CreateClientAccessServer(self, config, autoHandleToken=None):
        """Register new client access server.

        :param config: Configuration of new client access server
        :type config: str  (JSON)
        """
        return self.__callMethod("CreateClientAccessServer", {"config": config}, autoHandleToken=autoHandleToken)

    def CreateClientOneTimeUrlForArchiveAdmin(self, instanceID, instanceUrl=None, autoHandleToken=None):
        """Create URL including OTP for $archiveadmin access.

        :param instanceID:   Unique ID of MailStore instance in which this command should be invoked.
        :type instanceID:    str
        :param instanceUrl:  Base URL for accessing instance.
        :type instanceUrl:   str
        """

        return self.__callMethod("CreateClientOneTimeUrlForArchiveAdmin", {"instanceID": instanceID, "instanceUrl": instanceUrl}, autoHandleToken=autoHandleToken)

    def CreateDirectoryOnInstanceHost(self, serverName, path, autoHandleToken=None):
        """Create a directory on an Instance Host

        :param serverName:  Name of Instance Host.
        :type serverName:   str
        :param path:        Path of directory to create.
        :type path:         str
        this can be used to create empty directories for new instances"""
        return self.__callMethod("CreateDirectoryOnInstanceHost", {"serverName": serverName, "path": path}, autoHandleToken=autoHandleToken)

    def CreateInstance(self, config, autoHandleToken=None):
        """Creates new instance.

        :param config:
        :type config: str (JSON)
        a replacement method is available"""
        return self.__callMethod("CreateInstance", {"config": config}, autoHandleToken=autoHandleToken)

    def CreateInstanceHost(self, config, autoHandleToken=None):
        """Create a new Instance Host.

        :param config:  Configuration of new Instance Host.
        :type config:   str (JSON)
        """
        return self.__callMethod("CreateInstanceHost", {"config": config}, autoHandleToken=autoHandleToken)

    def CreateLicenseRequest(self, autoHandleToken=None):
        """Create and return data of a license request."""
        return self.__callMethod("CreateLicenseRequest", {}, autoHandleToken=autoHandleToken)

    def CreateProfile(self, instanceID, properties, raw="true", autoHandleToken=None):
        """Create a new archiving or exporting profile.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param properties:  Profile properties.
        :type properties:   str (JSON)
        :param raw:         Currently only 'true' is supported.
        :type raw:          bool
        """
        return self.__callMethod("CreateProfile", {"instanceID": instanceID, "properties": properties, "raw": raw}, autoHandleToken=autoHandleToken)

    def CreateStore(self, instanceID, name, path, requestedState=None, autoHandleToken=None):
        """Create and attach a new archive store.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param name:            Meaningful name of archive store.
        :type name:             str
        :param path:            Path of directory containing archive store data.
        :type path:             str
        :param requestedState:  State of archive store after attaching.
        :type requestedState    str
        """
        return self.__callMethod("CreateStore", {"instanceID": instanceID, "name": name, "path": path, "requestedState": requestedState}, autoHandleToken=autoHandleToken)

    def CreateSystemAdministrator(self, config, password, autoHandleToken=None):
        """Create a new SPE system administrator.

        :param config:    Configuration of new SPE system administrator.
        :type config:     str (JSON)
        :param password:  Password of new SPE system administrator.
        :type password:   str
        """
        return self.__callMethod("CreateSystemAdministrator", {"config": config, "password": password}, autoHandleToken=autoHandleToken)

    def CreateUser(self, instanceID, userName, privileges, fullName=None, distinguishedName=None, authentication=None, password=None, autoHandleToken=None):
        """Create new MailStore user.

        :param instanceID:         Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:          str
        :param userName:           User name of new MailStore user.
        :type userName:            str
        :param privileges:         Comma separated list of privileges.
        :type privileges:          str
        :param fullName:           Full name of user.
        :type fullName:            str
        :param distinguishedName:  LDAP DN string.
        :type distinguishedName:   str
        :param authentication:     Authentication setting for user: 'Standard' or 'DirectoryService'.
        :type authentication:      str
        :param password:           Password of new user.
        :type password:            str
        """
        if isinstance(privileges, (list, tuple)):
            privileges = ",".join(privileges)
        return self.__callMethod("CreateUser", {"instanceID": instanceID, "userName": userName, "privileges": privileges, "fullName": fullName, "distinguishedName": distinguishedName, "authentication": authentication, "password": password}, autoHandleToken=autoHandleToken)

    def DeleteClientAccessServer(self, serverName, autoHandleToken=None):
        """Delete Client Access Server from management database.

        :param serverName:  Name of Client Access Server.
        :type serverName:   str
        """
        return self.__callMethod("DeleteClientAccessServer", {"serverName": serverName}, autoHandleToken=autoHandleToken)

    def DeleteEmptyFolders(self, instanceID, folder=None, autoHandleToken=None):
        """Remove folders from folder tree that do not contain emails.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID    str
        :param folder:      Entry point in folder tree.
        :type folder:       str
        """
        return self.__callMethod("DeleteEmptyFolders", {"instanceID": instanceID, "folder": folder}, autoHandleToken=autoHandleToken)

    def DeleteInstanceHost(self, serverName, autoHandleToken=None):
        """Delete Instance Host from management database.

        :param serverName:  Name of Client Access Server.
        :type serverName:   str
        """
        return self.__callMethod("DeleteInstanceHost", {"serverName": serverName}, autoHandleToken=autoHandleToken)

    def DeleteInstances(self, instanceFilter, autoHandleToken=None):
        """Delete one or multiple MailStore Instances

        :param instanceFilter:  Instance filter string
        :type instanceFilter:   str
        """
        return self.__callMethod("DeleteInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def DeleteMessage(self, instanceID, id, autoHandleToken=None):
        """Delete a single message

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID    str
        :param id:          Unique ID of message. Format: <store_id>:<message_num>
        :type id:           str
        """
        return self.__callMethod("DeleteMessage", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def DeleteProfile(self, instanceID, id, autoHandleToken=None):
        """Delete an archiving or exporting profile.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of profile.
        :type id:           int
        """
        return self.__callMethod("DeleteProfile", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def DeleteSystemAdministrator(self, userName, autoHandleToken=None):
        """Delete SPE system administrator.

        :param userName:  User name of SPE system administrator.
        :type userName:  str
        """
        return self.__callMethod("DeleteSystemAdministrator", {"userName": userName}, autoHandleToken=autoHandleToken)

    def DeleteUser(self, instanceID, userName, autoHandleToken=None):
        """Delete a MailStore user.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param userName:    User name of MailStore user.
        :type userName:     str
        """
        return self.__callMethod("DeleteUser", {"instanceID": instanceID, "userName": userName}, autoHandleToken=autoHandleToken)

    def DetachStore(self, instanceID, id, autoHandleToken=None):
        """Detach archive store

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store.
        :type id:           int
        """
        return self.__callMethod("DetachStore", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def FreezeInstances(self, instanceFilter, autoHandleToken=None):
        """Freeze a MailStore Instance

        :param instanceFilter:  Instance filter string.
        :type instanceFilter:   str
        """
        return self.__callMethod("FreezeInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def GetArchiveAdminEnabled(self, instanceID, autoHandleToken=None):
        """Get current state of archive admin access.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetArchiveAdminEnabled", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetChildFolders(self, instanceID, folder=None, maxLevels=None, autoHandleToken=None):
        """Get child folders.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param folder:      Parent folder.
        :type folder:       str
        :param maxLevels:   Depth of child folders.
        :type maxLevels:    int
        """
        return self.__callMethod("GetChildFolders", {"instanceID": instanceID, "folder": folder, "maxLevels": maxLevels}, autoHandleToken=autoHandleToken)

    def GetClientAccessServers(self, withServiceStatus, serverNameFilter=None, autoHandleToken=None):
        """Get list of Client Access Servers.

        :param withServiceStatus:  Include service status or not.
        :type withServiceStatus:   bool
        :param serverNameFilter:   Server name filter string.
        :type serverNameFilter:    str
        """
        return self.__callMethod("GetClientAccessServers", {"serverNameFilter": serverNameFilter, "withServiceStatus": withServiceStatus}, autoHandleToken=autoHandleToken)

    def GetComplianceConfiguration(self, instanceID, autoHandleToken=None):
        """Get current compliance configuration settings.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID    str
        """
        return self.__callMethod("GetComplianceConfiguration", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetDirectoriesOnInstanceHost(self, serverName, path=None, autoHandleToken=None):
        """Get file system directory structure from Instance Host.

        :param serverName:  Name of Instance Host.
        :type serverName    str
        :param path:        Path of directory to obtain subdirectories from.
        :type path:         str
        """
        return self.__callMethod("GetDirectoriesOnInstanceHost", {"serverName": serverName, "path": path}, autoHandleToken=autoHandleToken)

    def GetDirectoryServicesConfiguration(self, instanceID, autoHandleToken=None):
        """Get current Directory Services configuration settings.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetDirectoryServicesConfiguration", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetEnvironmentInfo(self, autoHandleToken=None):
        """Return general information about SPE environment."""
        return self.__callMethod("GetEnvironmentInfo", {}, autoHandleToken=autoHandleToken)

    def GetFolderStatistics(self, instanceID, autoHandleToken=None):
        """Get folder statistics.

        :param instanceID: Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:  str
        """
        return self.__callMethod("GetFolderStatistics", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetIndexConfiguration(self, instanceID, autoHandleToken=None):
        """Get list of attachment file types to index.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetIndexConfiguration", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetInstanceConfiguration(self, instanceID, autoHandleToken=None):
        """Get configuration of MailStore Instance.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetInstanceConfiguration", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetInstanceHosts(self, serverNameFilter=None, autoHandleToken=None):
        """Get list of Instance Hosts.

        :param serverNameFilter:  Server name filter string.
        :type serverNameFilter:   str
        """
        return self.__callMethod("GetInstanceHosts", {"serverNameFilter": serverNameFilter}, autoHandleToken=autoHandleToken)

    def GetInstanceProcessLiveStatistics(self, instanceID, autoHandleToken=None):
        """Get live statistics from Instance process.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID    str
        """
        return self.__callMethod("GetInstanceProcessLiveStatistics", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetInstances(self, instanceFilter, autoHandleToken=None):
        """Get list of instances.

        :param instanceFilter:  Instance filter string.
        :type instanceFilter:   str
        """
        return self.__callMethod("GetInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def GetInstanceStatistics(self, instanceID, autoHandleToken=None):
        """Get archive statistics from instance.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetInstanceStatistics", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetMessages(self, instanceID, folder, autoHandleToken=None):
        """Get list of messages from a folder.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param folder:      Folder whose content to list.
        :type folder        str
        """
        return self.__callMethod("GetMessages", {"instanceID": instanceID, "folder" : folder}, autoHandleToken=autoHandleToken)

    def GetProfiles(self, instanceID, raw="true", autoHandleToken=None):
        """Get list of archiving and exporting profiles.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param raw:         Currently only 'true' is supported.
        :type raw:          bool
        """
        return self.__callMethod("GetProfiles", {"instanceID": instanceID, "raw": raw}, autoHandleToken=autoHandleToken)

    def GetServiceStatus(self, autoHandleToken=None):
        """Get current status of all SPE services."""
        return self.__callMethod("GetServiceStatus", {}, autoHandleToken=autoHandleToken)

    def GetStoreAutoCreateConfiguration(self, instanceID, autoHandleToken=None):
        """Get automatic archive store creation settings.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetStoreAutoCreateConfiguration", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetStores(self, instanceID, autoHandleToken=None):
        """Get list of archive stores.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetStores", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetSystemAdministrators(self, autoHandleToken=None):
        """Get list of system administrators."""
        return self.__callMethod("GetSystemAdministrators", {}, autoHandleToken=autoHandleToken)

    def GetTimeZones(self, instanceID, autoHandleToken=None):
        """Get list of available time zones.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("GetTimeZones", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetUserInfo(self, instanceID, userName, autoHandleToken=None):
        """Get detailed information about user.

        :param instanceID:   Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:    str
        :param userName:     User name of MailStore user
        :type userName:      str
        """
        return self.__callMethod("GetUserInfo", {"instanceID": instanceID, "userName": userName}, autoHandleToken=autoHandleToken)

    def GetUsers(self, instanceID, autoHandleToken=None):
        """Get list of users.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        """
        return self.__callMethod("GetUsers", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def GetWorkerResults(self, instanceID, fromIncluding, toExcluding, timeZoneID, profileID=None, userName=None, autoHandleToken=None):
        """Get results of profile executions.

        :param instanceID:     Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:      str
        :param fromIncluding:  Beginning of time range to fetch.
        :type fromIncluding:   str
        :param toExcluding:    End of time range to fetch.
        :type toExcluding:     str
        :param timeZoneID:     Time zone in which timestamp should be returned.
        :type timeZoneID:      str
        :param profileID:      Filter results by given profile ID.
        :type profileID:       str
        :param userName:       Filter results by given user name.
        :type userName:        str
        """
        return self.__callMethod("GetWorkerResults", {"instanceID": instanceID, "fromIncluding": fromIncluding, "toExcluding": toExcluding, "timeZoneID": timeZoneID, "profileID": profileID, "userName": userName}, autoHandleToken=autoHandleToken)

    def MaintainFileSystemDatabases(self, instanceID, autoHandleToken=None):
        """Execute maintenance task on archive store databases.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("MaintainFileSystemDatabases", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def MergeStore(self, instanceID, id, sourceId, autoHandleToken=None):
        """Merge two archive stores.

        :param instanceID: Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:  str
        :param id:         Unique ID of destination archive store.
        :type id           str
        :param sourceId:   Unique ID of source archive store.
        :type sourceId:    str
        """
        return self.__callMethod("MergeStore", {"instanceID": instanceID, "id" : id, "sourceId" : sourceId}, autoHandleToken=autoHandleToken)

    def MoveFolder(self, instanceID, fromFolder, toFolder, autoHandleToken=None):
        """Move folder.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param fromFolder:  Old folder name.
        :type fromFolder:   str
        :param toFolder:    New folder name.
        :type toFolder:     str
        """
        return self.__callMethod("MoveFolder", {"instanceID": instanceID, "fromFolder": fromFolder, "toFolder": toFolder}, autoHandleToken=autoHandleToken)

    def PairWithManagementServer(self, serverType, serverName, port, thumbprint, autoHandleToken=None):
        """Pair server role with Management Server.

        :param serverType:  Type of server role.
        :type serverType:   str
        :param serverName:  Name of server that hosts 'serverType' role.
        :type serverName:   str
        :param port:        TCP port on which 'serverType' role on 'serverName' accepts connections.
        :type port:         str
        :param thumbprint:  Thumbprint of SSL certificate used by serverType' role on 'serverName'.
        :type thumbprint:   str
        """
        return self.__callMethod("PairWithManagementServer", {"serverType": serverType, "serverName": serverName, "port": port, "thumbprint": thumbprint}, autoHandleToken=autoHandleToken)

    def Ping(self, autoHandleToken=None):
        """Send a keep alive packet."""
        return self.__callMethod("Ping", {}, autoHandleToken=autoHandleToken)

    def RebuildSelectedStoreIndexes(self, instanceID, autoHandleToken=None):
        """Rebuild search indexes of selected archive stores.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("RebuildSelectedStoreIndexes", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def RefreshAllStoreStatistics(self, instanceID, autoHandleToken=None):
        """Refresh archive store statistics.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("RefreshAllStoreStatistics", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def RenameStore(self, instanceID, id, name, autoHandleToken=None):
        """Rename archvive store

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store.
        :type id:           str
        :param name:        New name of archive store.
        :type id:           str
        """
        return self.__callMethod("RenameStore", {"instanceID": instanceID, "id": id, "name": name}, autoHandleToken=autoHandleToken)

    def RenameUser(self, instanceID, oldUserName, newUserName, autoHandleToken=None):
        """Rename a MailStore user.

        :param instanceID:   Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:    str
        :param oldUserName:  Old user name.
        :type oldUserName:   str
        :param newUserName:  New user name.
        :type newUserName:   str
        """
        return self.__callMethod("RenameUser", {"instanceID": instanceID, "oldUserName": oldUserName, "newUserName": newUserName}, autoHandleToken=autoHandleToken)

    def RestartInstances(self, instanceFilter, autoHandleToken=None):
        """Restart one or multiple instances.

        :param instanceFilter:  Instance filter string
        :type instanceFilter:   str
        """
        return self.__callMethod("RestartInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def RetryOpenStores(self, instanceID, autoHandleToken=None):
        """Retry opening stores that failed previously

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("RetryOpenStores", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def RunProfile(self, instanceID, id, autoHandleToken=None):
        """Run an existing archiving or exporting profile.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked."
        :type instanceID:   str
        :param id:          Unique profile ID.
        :type id:           str
        """
        return self.__callMethod("RunProfile", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def RunTemporaryProfile(self, instanceID, properties, raw="true", autoHandleToken=None):
        """Run a temporary/non-existent profile.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param properties:  Profile properties.
        :type properties:   str
        :param raw:         Currently only 'true' is supported.
        :type raw:          str
        """
        return self.__callMethod("RunTemporaryProfile", {"instanceID": instanceID, "properties": properties, "raw": raw}, autoHandleToken=autoHandleToken)

    def SelectAllStoreIndexesForRebuild(self, instanceID, autoHandleToken=None):
        """Select all archive store for rebuild.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        """
        return self.__callMethod("SelectAllStoreIndexesForRebuild", {"instanceID": instanceID}, autoHandleToken=autoHandleToken)

    def SetArchiveAdminEnabled(self, instanceID, enabled, autoHandleToken=None):
        """Enable or disable archive admin access.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param enabled:     Enable or disable flag.
        :type enabled:      bool
        """
        return self.__callMethod("SetArchiveAdminEnabled", {"instanceID": instanceID, "enabled": enabled}, autoHandleToken=autoHandleToken)

    def SetClientAccessServerConfiguration(self, config, autoHandleToken=None):
        """Set the configuration of a Client Access Server.

        :param config:  Client Access Server configuration.
        :type config:   str (JSON)
        """
        return self.__callMethod("SetClientAccessServerConfiguration", {"config": config}, autoHandleToken=autoHandleToken)

    def SetComplianceConfiguration(self, instanceID, config, autoHandleToken=None):
        """Set compliance configuration settings.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param config:      Compliance configuration.
        :type config:       str
        """
        return self.__callMethod("SetComplianceConfiguration", {"instanceID": instanceID, "config": config}, autoHandleToken=autoHandleToken)

    def SetDirectoryServicesConfiguration(self, instanceID, config, autoHandleToken=None):
        """Set directory services configuration settings.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param config:      Directory services configuration.
        :type config:       str
        """
        return self.__callMethod("SetDirectoryServicesConfiguration", {"instanceID": instanceID, "config": config}, autoHandleToken=autoHandleToken)

    def SetIndexConfiguration(self, instanceID, config, autoHandleToken=None):
        """Set full text search index configuration.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param config:      Full text search index configuration
        :type config        str (JSON)
        """
        return self.__callMethod("SetIndexConfiguration", {"instanceID": instanceID, "config": config}, autoHandleToken=autoHandleToken)

    def SetInstanceConfiguration(self, config, autoHandleToken=None):
        """Set configuration of MailStore Instance

        :param config:  Instance configuration.
        :type config:   str (JSON)
        """
        return self.__callMethod("SetInstanceConfiguration", {"config": config}, autoHandleToken=autoHandleToken)

    def SetInstanceHostConfiguration(self, config, autoHandleToken=None):
        """Set configuration of Instance Host.

        :param config:  Instance Host configuration.
        :type config:   str (JSON)
        """
        return self.__callMethod("SetInstanceHostConfiguration", {"config": config}, autoHandleToken=autoHandleToken)

    def SetStoreAutoCreateConfiguration(self, instanceID, config, autoHandleToken=None):
        """Set configuration for automatic archive store creation.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param config:      Archive store automatic creation configuration.
        :type config:       str (JSON)
        """
        return self.__callMethod("SetStoreAutoCreateConfiguration", {"instanceID": instanceID, "config": config}, autoHandleToken=autoHandleToken)

    def SetStorePath(self, instanceID, id, path, autoHandleToken=None):
        """Set the path to archive store data.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store.
        :type id:           int
        :param path:        Path to archive store data.
        :type path          str
        """
        return self.__callMethod("SetStorePath", {"instanceID": instanceID, "id": id, "path": path}, autoHandleToken=autoHandleToken)

    def SetStoreRequestedState(self, instanceID, id, requestedState, autoHandleToken=None):
        """Set state of archive store.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param id:              Unique ID of archive store.
        :type id:               int
        :param requestedState:  State ('normal','current','writeProtected','disabled')
        :type requestedState:   str
        """
        return self.__callMethod("SetStoreRequestedState", {"instanceID": instanceID, "id": id, "requestedState": requestedState}, autoHandleToken=autoHandleToken)

    def SetSystemAdministratorConfiguration(self, config, autoHandleToken=None):


        return self.__callMethod("SetSystemAdministratorConfiguration", {"config": config}, autoHandleToken=autoHandleToken)

    def SetSystemAdministratorPassword(self, userName, password, autoHandleToken=None):
        """Set password for SPE system administrator.

        :param userName:  User name of SPE system administrator.
        :type userName:   str
        :param password:  New password for SPE system administrator.
        :type password:   str
        """
        return self.__callMethod("SetSystemAdministratorPassword", {"userName": userName, "password": password}, autoHandleToken=autoHandleToken)

    def SetUserAuthentication(self, instanceID, userName, authentication, autoHandleToken=None):
        """Set authentication settings of a MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param authentication:  Authentication method. Either 'Standard' or 'Windows Authentication'.
        :type authentication:   str
        """
        return self.__callMethod("SetUserAuthentication", {"instanceID": instanceID, "userName": userName, "authentication": authentication}, autoHandleToken=autoHandleToken)

    def SetUserDistinguishedName(self, instanceID, userName, distinguishedName=None, autoHandleToken=None):
        """Set authentication settings of a MailStore user.

        :param instanceID:         Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:          str
        :param userName:           User name of MailStore user.
        :type userName:            str
        :param distinguishedName:  LDAP DN string.
        :type distinguishedName:   str
        """
        return self.__callMethod("SetUserDistinguishedName", {"instanceID": instanceID, "userName": userName, "distinguishedName": distinguishedName}, autoHandleToken=autoHandleToken)

    def SetUserEmailAddresses(self, instanceID, userName, emailAddresses=None, autoHandleToken=None):
        """Set email addresses of MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param emailAddresses:  List of email addresses.
        :type emailAddresses:   str
        """
        if isinstance(emailAddresses,  (list,tuple)):
            emailAddresses = ",".join(emailAddresses)
        return self.__callMethod("SetUserEmailAddresses", {"instanceID": instanceID, "userName": userName, "emailAddresses": emailAddresses}, autoHandleToken=autoHandleToken)

    def SetUserFullName(self, instanceID, userName, fullName=None, autoHandleToken=None):
        """Set full name of MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param fullName:        Full name of MailStore user.
        :type fullName:         str
        """
        return self.__callMethod("SetUserFullName", {"instanceID": instanceID, "userName": userName, "fullName": fullName}, autoHandleToken=autoHandleToken)

    def SetUserPassword(self, instanceID, userName, password, autoHandleToken=None):
        """Set password of MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param password:        Password of MailStore user.
        :type password:         str
        """
        return self.__callMethod("SetUserPassword", {"instanceID": instanceID, "userName": userName, "password": password}, autoHandleToken=autoHandleToken)

    def SetUserPop3UserNames(self, instanceID, userName, pop3UserNames=None, autoHandleToken=None):
        """Set POP3 user name of MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param pop3UserNames:   List of POP3 user names.
        """
        if isinstance(pop3UserNames, (list,tuple)):
            pop3UserNames = ",".join(pop3UserNames)
        return self.__callMethod("SetUserPop3UserNames", {"instanceID": instanceID, "userName": userName, "pop3UserNames": pop3UserNames}, autoHandleToken=autoHandleToken)

    def SetUserPrivileges(self, instanceID, userName, privileges, autoHandleToken=None):
        """Set privileges of MailStore user.

        :param instanceID:      Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:       str
        :param userName:        User name of MailStore user.
        :type userName:         str
        :param privileges:      Comma separated list of privileges.
        :type privileges:       str
        """
        if isinstance(privileges, (list, tuple)):
            privileges = ",".join(privileges)
        return self.__callMethod("SetUserPrivileges", {"instanceID": instanceID, "userName": userName, "privileges": privileges}, autoHandleToken=autoHandleToken)

    def SetUserPrivilegesOnFolder(self, instanceID, userName, folder, privileges, autoHandleToken=None):
        """Set privileges on folder for MailStore user.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param userName:    User name of MailStore user.
        :type userName:     str
        :param folder:      Folder name.
        :type folder:       str
        :param privileges:  Comma separated list of folder privileges.
        :type privileges:   str
        """
        if isinstance(privileges, (list, tuple)):
            privileges = ",".join(privileges)
        return self.__callMethod("SetUserPrivilegesOnFolder", {"instanceID": instanceID, "userName": userName, "folder": folder, "privileges": privileges}, autoHandleToken=autoHandleToken)

    def StartInstances(self, instanceFilter, autoHandleToken=None):
        """Start one or multiple MailStore Instances.

        :param instanceFilter:  Instance filter string
        :type instanceFilter:   str
        """
        return self.__callMethod("StartInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def StopInstances(self, instanceFilter, autoHandleToken=None):
        """Stop one or multiple MailStore Instances.

        :param instanceFilter:  Instance filter string
        :type instanceFilter:   str
        """
        return self.__callMethod("StopInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def SyncUsersWithDirectoryServices(self, instanceID, dryRun=None, autoHandleToken=None):
        """Sync users of MailStore instance with directory services.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param dryRun:      Simulate sync only.
        :type dryRun:       bool
        """
        if dryRun in ["True", "true", True, 1]:
            dryRun = "true"
        return self.__callMethod("SyncUsersWithDirectoryServices", {"instanceID": instanceID, "dryRun": dryRun}, autoHandleToken=autoHandleToken)

    def ThawInstances(self, instanceFilter, autoHandleToken=None):
        """Thaw one or multiple MailStore Instances.

        :param instanceFilter: Instance filter string.
        :type instanceFilter:  str
        """
        return self.__callMethod("ThawInstances", {"instanceFilter": instanceFilter}, autoHandleToken=autoHandleToken)

    def UpgradeStore(self, instanceID, id, autoHandleToken=None):
        """Upgrade archive store from MailStore Server 5 or older to current format.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store.
        :type id:           int
        """
        return self.__callMethod("UpgradeStore", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

    def VerifyStore(self, instanceID, id, autoHandleToken=None):
        """Verify archive stores consistency.

        :param instanceID:  Unique ID of MailStore instance in which this command is invoked.
        :type instanceID:   str
        :param id:          Unique ID of archive store.
        :type id:           int
        """
        return self.__callMethod("VerifyStore", {"instanceID": instanceID, "id": id}, autoHandleToken=autoHandleToken)

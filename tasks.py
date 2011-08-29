import gflags
import urllib
import httplib2
import time
import datetime


from apiclient.discovery import build
from oauth2client.file import Storage
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run
from rfc3339 import rfc3339

import pywintypes
import win32com.client

if win32com.client.gencache.is_readonly == True:
  #allow gencache to create the cached wrapper objects
  win32com.client.gencache.is_readonly = False

  # under p2exe the call in gencache to __init__() does not happen
  # so we use Rebuild() to force the creation of the gen_py folder
  win32com.client.gencache.Rebuild()

  # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
  # to allow creation of the cache in %temp%

PROXY_TYPE_HTTP = 3
PROXY_HOST = 'www-proxy.ericsson.se'
PROXY_PORT = 8080

GOOGLE_LIST_NAME = "Ericsson"

FLAGS = gflags.FLAGS


toOutlook = {'title' : 'Subject',  'notes' : 'Body', 'status' : 'Complete', 'id' : "EntryID"}
#  'due' : 'DueDate', 'updated' : 'LastModificationTime', 'completed' : 'DateCompleted'

# Important 
importantKeys = [  "Subject", "Complete", "Body", "EntryID", "LastModificationTime"]
# "ReminderTime", "CreationTime", "StartDate", "DueDate", "DateCompleted", "LastModificationTime",

toGoogle = dict ((v,k) for k,v in toOutlook.items())

def toDateTime(value):
  if value.year == 4501:
    # Combination of pywin being old and Outlook COM being stupid 
    # returns year 4501 if there is no due date 
    # (ie latest possible date acc'd to outlook)
    # Fix this to the max date rfc3339 will take?
    value  = rfc3339(datetime.datetime(2011,9,8,17,37,0))
    # value.year = 3000
    return [key,value]
    
  value = rfc3339(datetime.datetime(
    year=value.year,
    month=value.month,
    day=value.day,
    hour=value.hour,
    minute=value.minute,
    second=value.second
  ))
  return value

def toOutlookKey(item):
  key,value = item
  
  if key == "status":
    if value == "completed":
      value = 1 == 1
    else:
      value = not 1 == 1
  
  key = toOutlook[key]
  
  return [key,value]
  
  
def toGoogleKey(item):
  key,value = item
  key = toGoogle[key]

  if key == "status":
    if value:
      value = 'completed'
    else:
      value = 'needsAction'

  
  timeFields = {'updated','completed','due'}

  if key in timeFields:
    value = toDateTime(value)
    
  return [key,value]

def convertToGoogle(task):
  for key,value in task.items():
    if key not in toGoogle:
      del task[key]
  res = dict ((toGoogleKey(item)) for item in task.items())
  return res
  
def convertToOutlook(task):
  for key,value in task.items():
    if key not in toOutlook:
      del task[key]
  res = dict ((toOutlookKey(item)) for item in task.items())
  return res
  
  
class outlook:
  def __init__(self):
    self.records = []
    self.outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    # outlook = win32com.client.Dispatch("Outlook.Application")
    self.ns = self.outlook.GetNamespace("MAPI")
    ofTasks = self.ns.GetDefaultFolder(win32com.client.constants.olFolderTasks)

        
    for taskno in range(len(ofTasks.Items)):
      # print "taskno: ", taskno
      task = ofTasks.Items.Item(taskno+1)
      if task.Class == win32com.client.constants.olTask:
        keys = []
        # print "keys: ", len(task._prop_map_get_)
        #for key in task._prop_map_get_:
        # if isinstance(getattr(task,key), (int,str,unicode)):
        
        for key in task._prop_map_get_:
          
          if key in importantKeys:
            
            keys.append(key)
        
        record = {}
        for key in keys:
          record[key] = getattr(task,key)
        self.records.append(record)
      first = False
      
  def modify(self, task, taskid):
    updatetask = self.ns.GetItemFromID(taskid)
    for key,value in task.items():
      if not key == "EntryID":
        setattr(updatetask,key,value)
    updatetask.Save()
  
  def create(self, task):
    newtask = self.outlook.CreateItem(win32com.client.constants.olTaskItem)
    
    for key,value in task.items():
      # Set values for this new task, ensure EntryID isn't set
      if not key == "EntryID":
        setattr(newtask,key,value)
    newtask.Save()
    
    # Now convert this task into a dict format used elsewhere.
    # TODO: Combine this bit with the _init_ fuction...
    
    keys = []
    for key in newtask._prop_map_get_:
          
      if key in importantKeys:
        
        keys.append(key)
    record = {}
    for key in keys:
      record[key] = getattr(newtask,key)
    
    return record
    
    


class google:
  def __init__(self):
    # Set up a Flow object to be used if we need to authenticate. This
    # sample uses OAuth 2.0, and we set up the OAuth2WebServerFlow with
    # the information it needs to authenticate. Note that it is called
    # the Web Server Flow, but it can also handle the flow for native
    # applications
    # The client_id and client_secret are copied from the API Access tab on
    # the Google APIs Console
    FLOW = OAuth2WebServerFlow(
        client_id='45198696978.apps.googleusercontent.com',
        client_secret='PXAHwAr3i9vh13ckf2M89Zve',
        scope='https://www.googleapis.com/auth/tasks',
        user_agent='YOUR_APPLICATION_NAME/YOUR_APPLICATION_VERSION')

    # To disable the local server feature, uncomment the following line:
    # FLAGS.auth_local_webserver = False

    # If the Credentials don't exist or are invalid, run through the native client
    # flow. The Storage object will ensure that if successful the good
    # Credentials will get written back to a file.
    storage = Storage('tasks.dat')
    credentials = storage.get()



    if credentials is None or credentials.invalid == True:
      credentials = run(FLOW, storage)

    # Create an httplib2.Http object to handle our HTTP requests and authorize it
    # with our good Credentials.
    proxies = urllib.getproxies()
    # if len(proxies) > 0:
    if 1 < 2:
      # proxy_type, proxy_url = proxies.items()[0]
      # proxy_protocol, proxy_url = proxy_url.split('://')
      # proxy_url, proxy_port = proxy_url.split(':')
      # proxy_port = int(proxy_port)

    #temp until urllib works...
      proxy_type = PROXY_TYPE_HTTP
      proxy_url = PROXY_HOST
      proxy_port = PROXY_PORT

      http = httplib2.Http(proxy_info = httplib2.ProxyInfo(proxy_type, proxy_url, proxy_port),disable_ssl_certificate_validation=True)
      # http = httplib2.Http(proxy_info = httplib2.ProxyInfo(proxy_type, proxy_url, proxy_port))

    else:
      http = httplib2.Http(disable_ssl_certificate_validation=True)
    http = credentials.authorize(http)

    # Build a service object for interacting with the API. Visit
    # the Google APIs Console
    # to get a developerKey for your own application.
    self.service = build(serviceName='tasks', version='v1', http=http, developerKey='45198696978.apps.googleusercontent.com')

    self.tasklists = self.service.tasklists().list().execute()
    
  def modify(self,task,googlelistid,taskid):
    task['id'] = taskid
    return self.service.tasks().update(tasklist = googlelistid, body=task, task=taskid).execute()

  def update():
    self.tasklists = self.service.tasklists().list().execute()

  def add(self,task,googlelistid):
    # Need to strip the ID first
    del task['id']
    return self.service.tasks().insert(tasklist = googlelistid, body=task).execute()
    
  
def printa():
  for tasklist in googletasks.tasklists:
    print tasklist['title']

  for tasklist in outlooktasks.records:
    print tasklist["Subject"]

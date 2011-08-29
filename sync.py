
import pdb
import webbrowser
import datetime


import config
from tasks import *



conf = config.config()

googletasks = google()
outlooktasks = outlook()

# Find the outlook task list on google
for tasklist in googletasks.tasklists['items']:
  if GOOGLE_LIST_NAME == tasklist['title'] :
    result = tasklist
    googlelistid = result['id']
    break

# If the outlook task list doesn't exist on google then create it
else:
  tasklist = { 'title': GOOGLE_LIST_NAME }
  result = googletasks.service.tasklists().insert(body=tasklist).execute()
  googlelistid = result['id']
  googletasks.update
  

gtasks = googletasks.service.tasks().list(tasklist = result['id'] ).execute()


updatedG = 0
updatedO = 0
matched = 0
createdOnGoogle = 0
createdOnOutlook = 0

# Check if there are any tasks in google tasks
if 'items' in gtasks:
  # Add outlook tasks to google
  for otask in outlooktasks.records[:]:
   
    for gtask in gtasks['items'][:]:
      if gtask['id'] in conf.idMap and conf.idMap[gtask['id']] == otask['EntryID']:
        # ID's have been matched, update depending on modified times
        # Remove these tasks from otasks and gtasks as they don't need updating any more
        matched = matched + 1
        gtasks['items'].remove(gtask)
        outlooktasks.records.remove(otask)

        
        if otask['Subject'] != gtask['title']:
          # Something doesn't match, check modification times
          # print "Look at modified time of ", gtask['title']
          
          # gtime is always in utc
          gtime = datetime.datetime.strptime(gtask['updated'],"%Y-%m-%dT%H:%M:%S.%fZ")
          # otime is always in local time
          offset = datetime.datetime.now() - datetime.datetime.utcnow()
          otime = datetime.datetime.strptime(str(otask['LastModificationTime']),"%m/%d/%y %H:%M:%S") - offset
          
          if otime > gtime:
            # Replace google with outlook
            updatedG = updatedG + 1
            newtask = convertToGoogle(otask)
            gtask = googletasks.modify(newtask,googlelistid,conf.idMap[otask['EntryID']])
          else:
            # Replace outlook with google
            updatedO = updatedO + 1
            print "Replacing outlook with google"
            newtask = convertToOutlook(gtask)
            otask = outlooktasks.modify(newtask,conf.idMap[gtask['id']])
          
          
          
        elif 'notes' in gtask and otask['Body'] != gtask['notes']:
          # Something doesn't match, check modification times
          print "Look at modified time of ", gtask['title']
        break

    else:
      # doesn't exist, so add it
      
      createdOnGoogle = createdOnGoogle + 1
      newtask = convertToGoogle(otask)
      gtask = googletasks.add(newtask,googlelistid)
      conf.addMapping(otask,gtask)

  # Now need to add google tasks to outlook
  # Note that all matching ID's should have been done now, so we don't need to match ID's any more, just add them as we create them.
  # gtasks shoudl only contain tasks not in outlook now
  for gtask in gtasks['items']:
    
    createdOnOutlook = createdOnOutlook + 1
    otask = outlooktasks.create(convertToOutlook(gtask))
    conf.addMapping(otask,gtask)
    
else:
  # No tasks on Google tasks yet, Simply add all outlook tasks
  for otask in outlooktasks.records:
    createdOnGoogle = createdOnGoogle + 1
    newtask = convertToGoogle(otask)
    gtask = googletasks.add(newtask, googlelistid)
    conf.addMapping(otask,gtask)
  
print "Updated on Google: ",updatedG,"\nUpdated on Outlook: ",updatedO
print "Matched: ",matched,"\nCreated on Google: ",createdOnGoogle,"\nCreated on Outlook: ",createdOnOutlook

conf.dump()

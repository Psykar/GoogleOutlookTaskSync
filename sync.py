
import pdb
import webbrowser

import time
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
  
# Check if there are any tasks in google tasks
gtasks = googletasks.service.tasks().list(tasklist = result['id'] ).execute()


if 'items' in gtasks:
  
  
  # Add outlook tasks to google
  for otask in outlooktasks.records[:]:
   
    for gtask in gtasks['items'][:]:
      if gtask['id'] in conf.idMap and conf.idMap[gtask['id']] == otask['EntryID']:
        # ID's have been matched, update depending on modified times
        # Remove these tasks from otasks and gtasks as they don't need updating any more
        gtasks['items'].remove(gtask)
        outlooktasks.records.remove(otask)
        
        break
        """
        if 'notes' in gtask:
          if otask['Body'] == gtask['notes']:
            # Tasks match totally, note their ID's
            conf.addMapping(otask,gtask)
            print "-Totally matches!"
        elif otask['Body'] :
          # Need to add the body
          print "-otask has Body gtask doesn't??"
          #print otask['Subject']
          #print otask['Body']
          # print gtask['notes']
        else:
          print "Neither have a body!"
          
        
        """
        
    else:
      # doesn't exist, so add it
      print "!!!!!!!!!!!!!!!!doesn't exist!!!!!!!!!!!!!!!!!!!!"
      newtask = convertToGoogle(otask)
      gtask = googletasks.add(newtask,googlelistid)
      conf.addMapping(otask,gtask)

  # Now need to add google tasks to outlook
  # Note that all matching ID's should have been done now, so we don't need to match ID's any more, just add them as we create them.
  # gtasks shoudl only contain tasks not in outlook now
  for gtask in gtasks['items']:
    print "Creating outlook task..."
    otask = outlooktasks.create(convertToOutlook(gtask))
    conf.addMapping(otask,gtask)
    
else:
  # No tasks on Google tasks yet, Simply add all outlook tasks
  for otask in outlooktasks.records:
    newtask = convertToGoogle(otask)
    gtask = googletasks.add(newtask, googlelistid)
    conf.addMapping(otask,gtask)
  

conf.dump()

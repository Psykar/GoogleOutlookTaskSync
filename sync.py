
import pdb
import webbrowser
import datetime


import config
from tasks import *



conf = config.config()

googletasks = google(GOOGLE_LIST_NAME)
outlooktasks = outlook()

gtasks = googletasks.getTasks()
otasks = outlooktasks.getTasks()


updatedG = 0
updatedO = 0
matched = 0
createdOnGoogle = 0
createdOnOutlook = 0

def updateTask(task1, task2):
  # Something doesn't match, check modification times
  # create a new task to update the older side
  
  # TODO: Edit convert() to return a task object, not a dict
  # TODO: update the __eqs__ function for direct comaparison here
  global updatedG
  global updatedO
  time1 = task1.updatedUTC()
  time2 = task2.updatedUTC()
  
  if time1 > time2:
    newtask = task1.convert()
  else:
    newtask = task2.convert()
  
  if newtask.google:
    updatedG = updatedG + 1
    return googletasks.modify(newtask,conf.idMap[newtask['id']])
  elif newtask.outlook:
    updatedO = updatedO + 1
    return outlooktasks.modify(newtask,conf.idMap[newtask['id']])
  raise TypeError

# Add outlook tasks to google
for otask in otasks[:]:
  
  
  for gtask in gtasks[:]:
    if gtask['id'] in conf.idMap and conf.idMap[gtask['id']] == otask['id']:
      # ID's have been matched, update depending on modified times
      # Remove these tasks from otasks and gtasks as they don't need updating any more
      matched = matched + 1
      gtasks.remove(gtask)
      outlooktasks.tasks.remove(otask)
      
      if otask['title'] != gtask['title']:
        updateTask(otask,gtask)
      elif otask.completed() != gtask.completed():
        print otask['status'], gtask['status']
        updateTask(otask,gtask)
      elif 'notes' in gtask and otask['notes'] != gtask['notes']:
        updateTask(otask,gtask)
      break

  else:
    # doesn't exist, so add it
    createdOnGoogle = createdOnGoogle + 1
    newtask = otask.convertToGoogle()
    gtask = googletasks.add(newtask)
    
    conf.addMapping(otask,gtask)

# Now need to add google tasks to outlook
# Note that all matching ID's should have been done now, so we don't need to match ID's any more, just add them as we create them.
# gtasks shoudl only contain tasks not in outlook now
for gtask in gtasks:
  createdOnOutlook = createdOnOutlook + 1
  otask = outlooktasks.add(gtask.convertToOutlook())
  conf.addMapping(otask,gtask)
  
print "Updated on Google: ",updatedG,"\nUpdated on Outlook: ",updatedO
print "Matched: ",matched,"\nCreated on Google: ",createdOnGoogle,"\nCreated on Outlook: ",createdOnOutlook

conf.dump()

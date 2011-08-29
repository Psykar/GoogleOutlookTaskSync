
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


# Add outlook tasks to google
for otask in otasks[:]:
  print ".",
  
  for gtask in gtasks[:]:
    if gtask['id'] in conf.idMap and conf.idMap[gtask['id']] == otask['id']:
      # ID's have been matched, update depending on modified times
      # Remove these tasks from otasks and gtasks as they don't need updating any more
      matched = matched + 1
      gtasks.remove(gtask)
      outlooktasks.tasks.remove(otask)

      
      if otask['title'] != gtask['title']:
        # Something doesn't match, check modification times
        # print "Look at modified time of ", gtask['title']
        
        # gtime is always in utc
        # TODO: update the __eqs__ function for direct comaparison here
        gtime = datetime.datetime.strptime(gtask['updated'],"%Y-%m-%dT%H:%M:%S.%fZ")
        otime = otask.updatedUTC()
        
        if otime > gtime:
          updatedG = updatedG + 1
          # Replace google with outlook
          newtask = otask.convertToGoogle()
          gtask = googletasks.modify(newtask,conf.idMap[otask['id']])
        else:
          # Replace outlook with google
          updatedO = updatedO + 1
          newtask = gtask.convertToOutlook()
          otask = outlooktasks.modify(newtask,conf.idMap[gtask['id']])
        
        
        
      elif 'notes' in gtask and otask['notes'] != gtask['notes']:
        # Something doesn't match, check modification times
        print "Look at modified time of ", gtask['title']
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
  otask = outlooktasks.create(gtask.convertToOutlook())
  conf.addMapping(otask,gtask)
  
print "."
print "Updated on Google: ",updatedG,"\nUpdated on Outlook: ",updatedO
print "Matched: ",matched,"\nCreated on Google: ",createdOnGoogle,"\nCreated on Outlook: ",createdOnOutlook

conf.dump()

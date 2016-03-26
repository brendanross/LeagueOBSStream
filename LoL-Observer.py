from bs4 import BeautifulSoup
import requests
from eventlet.greenthread import sleep
import urllib
import os
import Image
import ImageGrab
import cv2
from cv2 import cv
from win32com.client import Dispatch
import win32api, win32con
import wmi
from eventlet.timeout import Timeout
from subprocess import call
scr = 'screen.png'
method = cv.CV_TM_SQDIFF_NORMED
SendKeys = Dispatch("WScript.Shell").SendKeys

matchDetail = "http://www.op.gg/summoner/ajax/spectator/"

def grabScreen():
	img = ImageGrab.grab()
	img.save("screen.png")
	
def compareImages(small):
	small = cv2.imread(small)
	large = cv2.imread(scr)
	result = cv2.matchTemplate(small, large, method)
	mn,_,mnLoc,_ = cv2.minMaxLoc(result)
	MPx,MPy = mnLoc
	return mn
	
def getMMR(team):
	totalMMR = 0
	for item in team:
		r = requests.Session()
		a = requests.adapters.HTTPAdapter(max_retries=5)
		r.mount("http://", a)
		page = r.get("http://www.op.gg"+item['href'])
		soup = BeautifulSoup(page.text)
		MMR = soup.findAll("span", {"class" : "mmr"})
		mmrIndex = MMR[0].text.index("MMR ")
		totalMMR = totalMMR + int(MMR[0].text[mmrIndex+4:])
	return totalMMR

def getSummoners():
	print "Searching for matches. Please wait.."
	r = requests.Session()
	a = requests.adapters.HTTPAdapter(max_retries=5)
	r.mount("http://", a)
	page = r.get('http://www.op.gg/spectate/pro/')
	soup = BeautifulSoup(page.text)
	matches = soup.findAll("div", {"class" : "SpectatorSummoner"})
	mmrValues = []
	spectateLinks = []
	i = 0
	if len(matches) == 0:
		return False
	
	while i < 6 and i < len(matches):
		print "Checking match " + str(i+1)
		summoner = matches[i].find("div", {"class" : "summonerName"})
		summoner = summoner.text.strip()
		# matchType = matches[i].find("div", {"class" : "QueueType"})
		# print len(matchType.text.strip())
		# if matchType.text != "Ranked Solo 5v5":
			# print "Not 5v5"
			# i=i+1
			# continue
		# else:
			# print "5v5"
		#preferrer summoners
		if "dade" in summoner or "SKT T1 Faker" in summoner or "hide on bush" in summoner or "SKT T1 Impact" in summoner:
			print "Preffered Summoner found."
			spectateButton = matches[i].find("div", {"class" : "SpectateButton"})
			spectateLink = spectateButton.a['href']
			postParams = {'userName' : summoner, 'force' : 'true'}
			r = requests.post(matchDetail, params=postParams)
			if r.status_code != 200:
				continue
			soup = BeautifulSoup(r.text)
			
			MMR = soup.findAll("td", {"class" : "Average"})
			
			if len(MMR) == 0:
				continue
			mmrIndex = MMR[0].text.index("MMR: ")
			teamOneMMR = int(MMR[0].text[mmrIndex+5:])
			teamTwoMMR = int(MMR[1].text[mmrIndex+5:])
			
			file = open('teamOne.txt', 'w')
			file.write(str(teamOneMMR))
			file.close()
			file = open('teamTwo.txt', 'w')
			file.write(str(teamTwoMMR))
			file.close()
			urllib.urlretrieve("http://www.op.gg"+spectateLink, "spectate.bat")
			break
		#end preferred
		
		spectateButton = matches[i].find("div", {"class" : "SpectateButton"})
		
		postParams = {'userName' : summoner, 'force' : 'true'}
		r = requests.post(matchDetail, params=postParams)
		if r.status_code != 200:
			i=i+1
			continue
		soup = BeautifulSoup(r.text)
		
		MMR = soup.findAll("td", {"class" : "Average"})
		
		if len(MMR) == 0:
			i=i+1
			continue
		mmrIndex = MMR[0].text.index("MMR: ")
		teamOneMMR = int(MMR[0].text[mmrIndex+5:])
		teamTwoMMR = int(MMR[1].text[mmrIndex+5:])
		matchTotalMMR = teamOneMMR+teamTwoMMR
		matchValues = [matchTotalMMR, teamOneMMR, teamTwoMMR]
		mmrValues.append(matchValues)
		
		spectateLinks.append(spectateButton.a['href'])
		sleep(3)
		i=i+1
	
	highest = 0
	for match in mmrValues:
		if match[0] > highest:
			highest = match[0]
	i = 0
	while i < len(mmrValues):
		if mmrValues[i][0] == highest:
			file = open('teamOne.txt', 'w')
			file.write(str(mmrValues[i][1]))
			file.close()
			file = open('teamTwo.txt', 'w')
			file.write(str(mmrValues[i][2]))
			file.close()
			urllib.urlretrieve("http://www.op.gg"+spectateLinks[i], "spectate.bat")
			break
		i=i+1
	return True

# def terminateProcess():
	# c = wmi.WMI()
	# for process in c.Win32_process():
		# if "League" in process.Name:
			# process.Terminate()
			# break

def monitorMatch():
	win = False
	while not win:
		sleep(2)
		grabScreen()
		kN = compareImages("EnglishVictory.png")
		eN = compareImages("KoreanVictory.png")
		if kN < 0.5 or eN < 0.5:
			SendKeys("2")
			SendKeys("%{F4}")
			#terminateProcess()
			win = True

#Basic DRM Functionality
def checkEnabled():
	r = requests.get('none')
	if r.text == "enable":
		return True
	else:
		return False

if __name__ == '__main__':
	print "Please keep League of Legends in the foreground while running."
	print "GameScene must be set to use 1 in OBS."
	print "BetweenMatchesScene must be set to 2 in OBS"
	print "Initializing.. "
	
	SendKeys("2")
	
	i = 0
	while True:
		foundMatch = False
		with Timeout(160, False) as timeout:
			foundMatch = getSummoners()
		
		if foundMatch:
			print "Starting match.."
			call("spectate.bat")
			
			sleep(8)
			SendKeys("1")
			print "Monitoring for match end.."
			monitorMatch()
			i=i+1
			print str(i)+" Matches Viewed Since Last Launch."
			if i == 2:
				os.system("start LoL-Stream.exe")
				#call("start LoL-Stream.exe")
				exit()
		else:
			sleep(10)
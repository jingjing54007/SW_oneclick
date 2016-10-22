'''
Created on Jun 30, 2015

@author: PGAUCHERAND
'''

import ConfigParser
import collections
import os, fnmatch
import shutil
import re
import colors as cons

#############################################################################################################	
# Custom exception class
#############################################################################################################
class ConfigParserException(Exception):
	def __init__(self,error):
		self.error=error
	def __str__(self):
		return repr(self.error)

#############################################################################################################	
# Parent class LogCompare
#############################################################################################################
class LogCompare(object):
	"""
	Compares the results of a local test with all the previous results on past modules/FW version.
	Class designed to be a common class with three target applications in mind: 
		- Standalone run with results printing on console
		- Run via Autotest with results printed on GUI
		- Run via one click without printing results
	"""
	def __init__(self,iniFile=None):
		"""
		@params iniFile : log compare configuration file 
		@var lstLogsToCompare : class attribute to store the list of logs that have been found on the system
		@var lstTests : class attribute to store the list of tests for which a log comparison is required
		@var testResults : class attribute to store the overall results of the comparison
		@raise IOError: if class is instanciated with a configuration file that doesn't exist
		"""
		self.lstLogsToCompare=collections.defaultdict(list)
		self.lstTests=[]
		self.testResults = collections.defaultdict(lambda: collections.defaultdict(list))
			
		if iniFile is not None:
			# if logCompare uses a config file then parse it
			if os.path.exists(iniFile):
				self.cfgFile=iniFile
				# parse the config file
				self.cDict=self.cfgDict()
			else:
				raise IOError("Config file provided doesn't exist in the file system : %s couldn't be found" %iniFile)
			
	#############################################################################################################	
	# Core functions
	#############################################################################################################
	
	def cfgDict(self):
		"""
		Parses the configuration file and store the data structure in a dictionary.
		@return dictionary containing the different elements of the configuration file
		"""
		cfgDict=collections.defaultdict(dict)
		Config = ConfigParser.ConfigParser()
		Config.read(self.cfgFile)
		sections = Config.sections()
		for section in sections:
			options = Config.options(section)
			for option in options:
				cfgDict[section][option]=Config.get(section,option)
		return cfgDict

	def validateInput(self):
		"""
		Validates the input data and if any problem is found raises ConfigParserException
		"""
		# Validate Paths
		for key in self.cDict['Paths'].keys():
			if key not in ['local_log','repo_log']:
				raise ConfigParserException("Wrong keys used in Paths section. Can only be 'local_log','repo_log'")
				
		if not os.path.exists(self.cDict['Paths']['local_log']):
			raise ConfigParserException("Path to local logs doesn't exist. Please double check configuration file and make sure the path exists")
		
		if not os.path.exists(self.cDict['Paths']['repo_log']):
			raise ConfigParserException("Path to repo logs doesn't exist. Please double check configuration file and make sure the path exists")
		
		
		# Validate list of tests - no need for validation yet
		
		# Validate info
		for key in self.cDict['Info'].keys():
			if key not in ['mod','fw']:
				raise ConfigParserException("Wrong keys used in Paths section. Can only be 'mod','fw'")
		if self.cDict['Info']['mod']=="" or self.cDict['Info']['fw']=="":
			raise ConfigParserException("Please enter value for mod and fw (typically is the locally tested module and FW)")
		
		# Validate output
		for key in self.cDict['Results'].keys():
			if key not in ['results_path']:
				raise ConfigParserException("Wrong keys used in Paths section. Can only be 'results_path'")
			
	def findLstLogsToCompare(self):
		"""
		Searches the list of logs to compare in the common log storage location.
		@precondition: path to perform the log search exists
		@postcondition: class attribute lstLogsToCompare is populated
		@return dictionary containing the list of every path to every log file per test
		@raise IOError if no logs are found in location
		"""
		for root, dirs, files in os.walk(self.cDict['Paths']['repo_log']):
			for name in files:
				for test_name in self.lstTests:
					if fnmatch.fnmatch(name.split('.')[0], test_name):
						self.lstLogsToCompare[test_name].append(os.path.join(root, name))
		
		if not self.lstLogsToCompare:
			raise IOError("Sorry, no logs have been found : \nSearch location : %s\nFiles searched : %s" %(self.cDict['Paths']['repo_log'],self.lstTests))

	def lookForMissingLogs(self):
		"""
		Searches for missing logs for every test for which a log compare is to be performed
		@precondition: class attribute lstLogsToCompare is populated
		"""	
		if len(self.lstLogsToCompare.keys()) != len(self.lstTests):
			for item in self.lstTests:
				if item not in self.lstLogsToCompare.keys():
					raise IOError("Couldn't find test for the following test : %s in %s" %(item,self.cDict['Paths']['repo_log']))
						
	def copyLogsLocally(self):
		"""
		Copies all the log files locally
		Avoids processing data on network paths.
		@returns dictionary of path to local log file for every test
		@raise Exception if the copy fails
		"""
		localCompareRepo=collections.defaultdict(list)
		# print self.lstLogsToCompare
		for key in self.lstLogsToCompare:
			lstDest=[]
			for item in self.lstLogsToCompare[key]:
				directory=os.path.dirname(item)
				drive_letter=directory.split("\\")[0]
				lstDest.append(directory.replace(drive_letter,"C:"))
			try:
				for i in range(0,len(lstDest)):
					if not os.path.exists(lstDest[i]):
						os.makedirs(lstDest[i])
					shutil.copy(self.lstLogsToCompare[key][i],lstDest[i])
				localCompareRepo[key]=lstDest
			except Exception:
				raise
		return localCompareRepo
		
	def compareIndividualLog(self,repo_log, local_log_path,test_name,ModFW):
		"""
		Individual log comparison. Compares two log files
		Ignores the preparation phase of a test and only retrieves errors encountered in the test.
		
		@params repo_log : path to previous' run log file
		@params local_log_path : path to local run's log file
		@params test_name : test name 
		@params ModFW : module and FW for the previous' run log file
		
		@return 0 the local run fails for the same reason compared to previous run 
		@return 1 the local run fails for a different reason compared to previous run
		@return 2 the previous' run passed
		@return 3 the local run fails for a different reason compared to previous run but the number of errors is different so a manual comparison is required
		"""
		
		tLocal=[]
		tRepo=[]
		
		# flag to enable comparing just the useful part of the test. I.e test set up doesn't intervene in comparison
		# this way, we can compare tests that have started with module already ON and tests that started from scratch
		cmp_flag_local=0
		cmp_flag_repo=0
		cmp_flag_local_failed=False
		cmp_flag_local_passed=False
		cmp_flag_repo_passed=False
		cmp_flag_repo_failed=False
		
		try:
			local_log_file = open(os.path.join(local_log_path,test_name+".log"))
			repo_log_file  = open(os.path.join(repo_log,test_name+".log"))

			# get the errors from the local run in local log file
			for line_local in local_log_file.readlines():
				if re.match('.*----- Testing Start -----.*',line_local): 
					cmp_flag_local=1
				if re.match('^Expected.Response.*|^Received.Response.*',line_local) and cmp_flag_local==1 : 
					tLocal.append(line_local)
				elif re.match('Status.*%s.*: PASSED' %test_name,line_local):
					cmp_flag_local_passed=True
				elif re.match('Status.*%s.*: FAILED' %test_name,line_local):
					cmp_flag_local_failed=True
			
			# get the errors from the repo run in repo folder
			for line_repo in repo_log_file.readlines():
				if re.match('.*----- Testing Start -----.*',line_repo): 
					cmp_flag_repo=1
				if re.match('^Expected.Response.*|^Received.Response.*',line_repo) and cmp_flag_repo==1 : 
					tRepo.append(line_repo)
				elif re.match('Status.*%s.*: PASSED' %test_name,line_repo):
					cmp_flag_repo_passed=True
				elif re.match('Status.*%s.*: FAILED' %test_name,line_repo):
					cmp_flag_repo_failed=True

			local_log_file.close()
			repo_log_file.close()
		
			# no errors have been found in previous run logs
			if not tRepo:
				# previous run results is passed and current run result is failed
				if cmp_flag_repo_passed and cmp_flag_local_failed:
					return 2
				# previous run results is failed and current run result is failed
				elif cmp_flag_repo_failed and cmp_flag_local_failed:
					return 4
				# previous run results is passed and current run result is passed
				elif cmp_flag_repo_passed and cmp_flag_local_passed:
					return 0
			# no errors have been found in current run logs but some errors in previous run logs
			elif not tLocal and tRepo:
				# previous run is failed and local log is passed
				if cmp_flag_repo_failed and cmp_flag_local_passed:
					return 5
				# undetermined behavior
				else:
					return 3
			else:
				# same number of errors found in both current and previous run
				if len(tLocal)==len(tRepo):
					cmpList = [[i,j] for i, j in zip(tLocal,tRepo) if i != j]
					if not cmpList:
						#if this list is empty then the logs have exactly the same content
						return 0
					else:
						# Temporary writing of the result in a file
						# Should think of a more generic way of doing it
						#self.writeCmpResult(cmpList,test_name,ModFW[0]+"_"+ModFW[1])
						return 1
				# different number of errors between the current run and previous run
				else:
					return 3
				
		except Exception:
			raise
		
	def compareLogs(self,lstLogs, local_log_path,lstTests):
		"""
		Compares multiple log files in a sequential manner for multiple tests
		
		Populates the following class attributes (to allow reusability in another program)
		Success contains the mod/FW for which exhibits same failure as local run
		Fail contains the mod/FW for which exhibits different failure compared with local run
		Pass contains the mod/FW for which exhibits the test has passed in the past
		TBA contains the mod/FW for which exhibits different failure compared with local run and for which a manual check is required
		
		@params lstLogs : dictionnary of path to logs files for different tests
		@params local_log_path : path to local run logs location
		@params lstTests : list of tests for which is comparison is performed
		"""
		
		for test_name in lstTests:
			lstSuccess=[]
			lstFail=[]
			lstPass=[]
			lstTBA=[]
			lstScriptUpdNeed=[]
			lstItemCoverage=[]
			
			for item in lstLogs[test_name]:
				ModFW=self.extractModFwVersion(item)
				#HL7528 has a one click folder without any SW version which is wrong. So ignore that because we can't know the SW version
				if ModFW[0]==None or ModFW[1]==None:
					continue
				rCmp = self.compareIndividualLog(item,local_log_path,test_name,ModFW)
				if rCmp==0:
					lstSuccess.append(ModFW)
				elif rCmp==1:
					lstFail.append(ModFW)
				elif rCmp==2:
					lstPass.append(ModFW)
				elif rCmp==3:
					lstTBA.append(ModFW)
				elif rCmp==4:
					lstScriptUpdNeed.append(ModFW)
				elif rCmp==5:
					lstItemCoverage.append(ModFW)
			
			self.testResults[test_name]['Success']=self.removeDictDuplicates(self._2DlstToDict(lstSuccess))
			self.testResults[test_name]['Fail']=self.removeDictDuplicates(self._2DlstToDict(lstFail))
			self.testResults[test_name]['Pass']=self.removeDictDuplicates(self._2DlstToDict(lstPass))
			self.testResults[test_name]['TBA']=self.removeDictDuplicates(self._2DlstToDict(lstTBA))
			self.testResults[test_name]['ScriptUpdNeed']=self.removeDictDuplicates(self._2DlstToDict(lstScriptUpdNeed))
			self.testResults[test_name]['ItemCoverage']=self.removeDictDuplicates(self._2DlstToDict(lstItemCoverage))
			
	def runCompare(self):
		"""
		Finds the list of logs to compare
		Searches for missing logs
		Copies all the file locally for future processing
		Proceeds to a one to one comparison for each test that need to be compared
		"""
		#dict containing the location of all the logs files found per test
		self.findLstLogsToCompare()
		
		self.lookForMissingLogs()
		
		# copy is done for all the tests at once in order to do os.walk on the network drive only once
		# new list of local path is edited and added to localRepo attribute
		localRepo=self.copyLogsLocally()
		
		# compare logs
		self.compareLogs(localRepo,self.cDict['Paths']['local_log'],self.lstTests)
	
	def _2DlstToDict(self,lst):
		"""
		Converts a 2D list into a dictionnary
		"""
		resDict=collections.defaultdict(list)
		for item in lst:
			resDict[item[0]].append(item[1])
		return resDict
	
	def removeDictDuplicates(self,resDict):
		"""
		Remove duplicates from dictionary values
		"""
		newDict=collections.defaultdict(list)
		for key in resDict.keys():
			tempLst = list(set(resDict[key]))
			tempLst.sort()
			newDict[key] = tempLst
		return newDict
		
	def extractModFwVersion(self,repo_log):
		"""
		Extracts the module type and FW version from a log file path.
		For example : Y:\RD_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\Intel\HL7528\BHL7528.2.9.152000.201506181422.x7160_1\Autotest
		Mod = HL7528
		FW = BHL7528.2.9.152000.201506181422.x7160_1
		"""
		Mod=None
		FW=None
		split=repo_log.split("\\")
		# The benefit from using groups is that if the path changes in the future
		# the mod and FW will still be able to be retrieved without changing this code
		patternModule=re.compile('^HL.*')
		patternFW = re.compile('(.*\.){3,10}')
		for item in split:
			if re.match(patternModule,item):
				Mod=item
			elif re.match(patternFW,item):
				FW=item
		return [Mod,FW]
	
	#############################################################################################################	
	# Print functions
	#############################################################################################################
		
	def customOutput(self,text,outpuType):
		"""
		Changes print colour on manual to have more readable results
		@params text : text to print on manual
		@params outpuType : colour with which text is printed
		"""
		default_colors = cons.get_text_attr()
		default_bg = default_colors & 0x0070
		
		
		if outpuType=="HEADER":
			cons.set_text_attr(cons.FOREGROUND_BLUE | default_bg | cons.FOREGROUND_INTENSITY)
		if outpuType=="FAIL":
			cons.set_text_attr(cons.FOREGROUND_RED | default_bg | cons.FOREGROUND_INTENSITY)
		if outpuType=="WARNING":
			cons.set_text_attr(cons.FOREGROUND_YELLOW | default_bg | cons.FOREGROUND_INTENSITY)
		if outpuType=="PASS":
			cons.set_text_attr(cons.FOREGROUND_GREEN | default_bg | cons.FOREGROUND_INTENSITY)
	
		print text
		
	def printResult(self,test_name):
		"""
		Prints comparison results to manual
		@params test_name : test for which the log comparison results are displayed
		"""
		self.customOutput("\n#################################################################","HEADER")
		self.customOutput("Test case : %s" %test_name,"HEADER")
		self.customOutput("#################################################################\n","HEADER")
		self.customOutput("Locally run version : %s : %s\n" %(self.cDict['Info']['mod'],self.cDict['Info']['fw']),"HEADER")
		
		if self.testResults[test_name]['Success']:
			self.customOutput("Same failure on :","PASS")
			self.printModFwDict(self.testResults[test_name]['Success'],"PASS")
			# self.writeOKResult(test_name,self.testResults[test_name]['Success'])

		if self.testResults[test_name]['Fail']:
			self.customOutput("\nDifferent failure on :","WARNING")
			self.printModFwDict(self.testResults[test_name]['Fail'],"WARNING")
		
		if self.testResults[test_name]['Pass']:
			self.customOutput("\nTest passed on :","PASS")
			self.printModFwDict(self.testResults[test_name]['Pass'],"PASS")

		if self.testResults[test_name]['TBA']:
			self.customOutput("\nDifferent failure with a different number of error :","FAIL")
			self.printModFwDict(self.testResults[test_name]['TBA'],"FAIL")
	
	def printModFwDict(self,aDict,outpuType):
		"""
		Prints module and FW from aDict to manual
		@params aDict : dictionnary of module and FW extracted from log paths
		@params outpuType  : colour with which text is printed
		"""
		for key in aDict.keys():
			self.customOutput("================================================================",outpuType)
			self.customOutput("=    %s    ==    %s" %(key,aDict[key][0]),outpuType)
			for item in aDict[key][1:]:
				self.customOutput("=              ==    %s" %item,outpuType)
		self.customOutput("================================================================",outpuType)
			
	def writeCmpResult(self,diffBtwLogs,test_name, repo_version):
		"""
		Write differences found in logs between two runs to a local file (configured in configuration file)
		@params diffBtwLogs : list of differences found in logs between two runs
		@params test_name   : name of the test
		@params repo_version: module and FW version of previous run
		"""
		resPath=os.path.join(self.cDict['Results']['results_path'],test_name.split(".")[0])
		local_version = self.cDict['Info']['mod']+"_" + self.cDict['Info']['fw']
		resFile = os.path.join(resPath,repo_version+".log")
		#in order to have a readable display in the logs, both version names should have the same length
		if len(local_version)!=len(repo_version):
			lDiff = abs(len(local_version)-len(repo_version))
			for i in range(0,lDiff):
				if len(local_version)<len(repo_version):
					local_version+=" "
				elif len(local_version)>len(repo_version):
					repo_version+=" "
		if not os.path.exists(resPath):
			os.makedirs(resPath)
		with open(resFile,'w') as logFile:
			logFile.write("Differences found between %s and %s\n" %(local_version,repo_version))
			for diff in diffBtwLogs:
				logFile.write("\n###############################################################\n")
				logFile.write("%s log : %s\n" %(local_version,diff[0].strip("\n")))
				logFile.write("%s log : %s\n" %(repo_version,diff[1].strip("\n")))
		logFile.close()

	def writeOKResult(self,test_name,lTestSuccess):
		"""
		Writes the successful comparison on manual and log file
		@params test_name    : name of the test 
		@params lTestSuccess : list of the modules and FW with the same results
		"""
		local_version = self.cDict['Info']['mod']+"_" + self.cDict['Info']['fw']
		resPath=os.path.join(self.cDict['Results']['results_path'],test_name.split(".")[0])
		resFile = os.path.join(resPath,"ModFW_with_matching_behaviour.log")
		if not os.path.exists(resPath):
			os.makedirs(resPath)
		if os.path.exists(resFile):
			os.remove(resFile)
		
		#only write results if comparison has been successful (i.e : lTestSuccess is not empty)
		if lTestSuccess:			
			with open(resFile,'a') as logFile:
				# logFile.write("################################################################################\n")
				# self.customOutput("\n################################################################################","HEADER")
				# logFile.write("Test case : %s\n" %test_name)
				# self.customOutput("Test case : %s\n" %test_name,"HEADER")
				logFile.write("################################################################################\n")
				logFile.write("The following Module/FW have the same behaviour as %s\n" %local_version)
				
				for key in self.testResults[test_name]['Success'].keys():
					logFile.write("\n==>Module %s\n" %key)
					for item in self.testResults[test_name]['Success'][key]:
						logFile.write("====> for FW : %s\n" %item)
			logFile.close()
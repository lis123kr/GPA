"""
 Copyright [2017] [Il Seop Lee / Chungbuk National University]
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 """

from logging import error, info
from datetime import datetime

def infolog(message):
	info("{0} {1}".format(datetime.now().isoformat(), message))
	print("{0} {1}".format(datetime.now().isoformat(), message))

def errorlog(message):
	error("{0} {1}".format(datetime.now().isoformat(), message))	
	print("{0} {1}".format(datetime.now().isoformat(), message))
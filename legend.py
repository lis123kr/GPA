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
class Legend:

	def __init__(self, fig, config):
		self.fig = fig

		self.config = config

		self.leg = self.fig.legend(handles=config['spatches'], labels=config['label'], bbox_to_anchor=config['bbox'], loc=config['loc'],
			ncol=config['ncol'], mode=config['mode'], borderaxespad=config['borderaxespad'])

		for legpatch in self.leg.get_patches():
			legpatch.set_picker(7)

		self.onpick = None

	def pick(self):
		from matplotlib.backend_bases import MouseEvent
		for lb, legpatch in zip(self.config['label'],self.leg.get_patches()):
			legpatch.pick(MouseEvent(lb, self.fig.canvas, 0, 0))

	def get_legend(self):
		return self.leg
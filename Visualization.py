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
import matplotlib.pyplot as plt
from numpy import where, linspace, logical_and, logical_or, array, abs, logical_not
from mcolors import mcols, scols
from Slider import ZoomSlider
# from matplotlib.widgets import Button
class visualization:
	def __init__(self, book):
		# from pandas import ExcelFile, DataFrame, merge
		from VizBook import vizbook

		self.duma_position = book.col_DumaPosition
		self.duma_Genome = book.col_GenomeStructure
		self.duma_Repeat = book.col_RepeatRegion
		self.duma_ORF = book.col_ORF

		self.dumas_col = [ self.duma_position, self.duma_Genome, self.duma_Repeat, self.duma_ORF ]
		self.dumas = { self.duma_Genome : [], self.duma_Repeat: [], self.duma_ORF : [] }
		self.spans = dict() #{ duma_Genome : dict(), duma_Repeat : dict(), duma_ORF : dict()}

		self.seq_size, self.markersize = 100, 10
		self.xlin = linspace(0.0, 1.0, self.seq_size)
		self.mcolor = [ 'Blues', 'Orgs', 'Grns', 'Reds' ]
		self.dcolor = [ [None, 'BuOg', 'BuGn', 'BuRd' ], 
					 ['OgBu', None, 'OgGn', 'OgRd'], 
				 	 ['GnBu', 'GnOg', None, 'GnRd'], 
				 	 ['RdBu', 'RdOg', 'RdGn', None] ]
		self.basecolors = ['blue', 'orange', 'lightgreen', 'red']
		self.basepair = ['A', 'G', 'C', 'T']
		self.annot, self.sct, self.leg = [], [], None
		self.annot_gere = dict()
		self.B = None
		self.bpdiff = None
		self.major_visiable = [ True, True, True, True]
		self.base_range, self.bidx = [ 0., 5., 10., 15., 20., 25., 30., 35., 40., 45.], 0
		self.variation_range, self.vidx = [ 5., 10., 15., 20., 25., 30., 35., 40., 45.], 0

		self.B = vizbook(book)
		self.B.preprocessing()
		self.dumas = self.B.dumas_info
		self.nsheet = self.B.nsheet
		self.bpdiff = self.B.bpdiff
		self.ylabel = self.B.datalabel

		# 임의지정을 통한 교환
		self.ylabel[1], self.ylabel[2] = self.ylabel[2], self.ylabel[1]
		self.bpdiff[1], self.bpdiff[2] = self.bpdiff[2], self.bpdiff[1]
		# bpdiff_trans[1], bpdiff_trans[2] = bpdiff_trans[2], bpdiff_trans[1]
	

	def visualize(self):
		import matplotlib.patches as mpatches

		self.fig, self.ax = plt.subplots(3, 1, sharex=True)
		self.ax[0].set_title('Click on legend rectangle to toggle data on/off', fontdict={'fontsize' : 8 })
		self.fig.set_size_inches(12, 7, forward=True)
		plt.subplots_adjust(left=0.05, bottom=0.15, right=0.98, hspace=0.15)
		# bpdiff_trans = [] # major transition

		self.ax = self.ax[::-1]

		# major가 같은 것 들만
		print('Run matplotlib...')
		for i in range(len(self.bpdiff)):
			sc = [ {'major' : [], 'minor': [] }, {'major' : [], 'minor': [] }, 
			       {'major' : [], 'minor': [] }, {'major' : [], 'minor': [] } ]

			self.draw_vspans(self.ax[i], self.B.range_dumas)
			
			for a in range(4):

				cond1 = self.bpdiff[i]['Mj_seq_x'] == a

				# major a -> minor b
				for b in range(4):
					if a == b:
						# sc[a]['major'].append(None)
						# sc[a]['minor'].append(None)
						continue

					# 1. minor가 같은 것들 처리 : 같은 색상
					cond2 = self.bpdiff[i]['Mn_seq_y'] == b
					# x는 
					cond3 = logical_or(self.bpdiff[i]['minor_x'] == 0, self.bpdiff[i]['Mn_seq_x'] == b)				
					cond3 = logical_and(cond2, cond3)

					d = self.B.get_bpdiff_cond(index=i, cond=logical_and(cond1, cond3)) #bpdiff[i][ logical_and(cond1, cond3).values ]
					if len(d) > 0:
						y_ = [ linspace(d_.maf_x, d_.maf_y, self.seq_size) for d_ in d[['maf_x', 'maf_y']].itertuples() ]					
						x_ = array(y_)
						for j in range(len(d)):
							x_[j].fill( d[self.duma_position].values[j] )
						c = [ self.xlin for _ in range(len(x_))]
						sc[a]['minor'].append(self.ax[i].scatter(x_, y_, c=c, cmap=mcols.cmap(self.mcolor[b]), s=self.markersize, linewidth=0.0))

						# draw major scatter
						y_ = self.get_uppery(y_)
						x_ = [ ii[0] for ii in x_]
						sc[a]['major'].append(self.ax[i].scatter(x_, y_, color=self.basecolors[a], label=self.basepair[a], s=self.markersize-1))
					
					# # 2. minor가 바뀌는 것들 처리 : colors diverging
					# 각 basetype -> b type, target이 b, 즉, b로 변한것
					cond2 = logical_and(self.bpdiff[i]['Mn_seq_x'] != b, self.bpdiff[i]['minor_x'] > 0)				
					cond3 = logical_and(cond2, self.bpdiff[i]['Mn_seq_y'] == b)
					d = self.bpdiff[i][ logical_and(cond1, cond3).values ]
					if len(d) > 0:
						self.draw_minor_transition(self.ax[i], sc, d, a, b)
				# end for b
			#end for a
			self.sct.append(sc)

			self.ax[i].set_ylim([0, 50])
			self.ax[i].set_ylabel('Variation of MAF(%)')
			self.ax[i].set_title(self.ylabel[i], loc='left', fontdict={'fontsize':13, 'verticalalignment':'top', 'color':'black', 'backgroundcolor':'#FEFEFE'})

		# ZoomSlider of Dumas position
		zaxes = plt.axes([0.08, 0.07, 0.90, 0.03], facecolor='lightgoldenrodyellow')
		self.zslider = ZoomSlider(zaxes, 'Dumas Position', valmin=self.B.minpos, valmax=self.B.maxpos, valleft=self.B.minpos, valright=self.B.maxpos, color='lightblue', valstep=1.0)
		self.zslider.on_changed(self.update_xlim)


		from legend import Legend
		spatches = [ mpatches.Patch(color=self.basecolors[idx], label=self.basepair[idx]) for idx in range(4) ]
		leg_basepair = Legend(self.fig, { 'spatches' : spatches, 'label' : self.basepair, 'bbox':(0.66, .91, 0.33, .102), 'loc' : 3,
					'ncol' : 4, 'mode' : 'expand', 'borderaxespad' : 0.})

		dumalabels = [ key for key in self.dumas.keys()]
		spatches = [ mpatches.Patch(color='grey', label=dumalabels[idx]) for idx in range(3) ]
		leg_dumas = Legend(self.fig, { 'spatches' : spatches, 'label' : dumalabels, 'bbox':(0.05, .91, 0.33, .102), 'loc' : 3,
					'ncol' : 3, 'mode' : 'expand', 'borderaxespad' : 0.})	

		spatches = [ mpatches.Patch(color=['black', 'magenta'][idx], label=['base', 'variation'][idx]) for idx in range(2) ]
		self.leg_filter = Legend(self.fig, { 'spatches' : spatches, 'label' : ['base_0.0', 'variation_5.0'], 'bbox':(0.40, .91, 0.24, .102), 'loc' : 3,
					'ncol' : 3, 'mode' : 'expand', 'borderaxespad' : 0.})	

		for x in self.ax:
			self.annot.append(x.annotate("", xy=(0,0), xytext=(20,-30),textcoords="offset points",
	                    bbox=dict(boxstyle="round", fc="w"), arrowprops=dict(arrowstyle="->")))
			# annotate over other axes
			x.figure.texts.append(x.texts.pop())

			# 
			# x.callbacks.connect('xlim_changed', xlim_changed_event)

		for c in self.B.range_dumas.keys():
			self.annot_gere[c] = []
			ay_ = 0.85 if c == self.duma_Repeat else 0.7
			if c == self.duma_ORF: ay_ = 1.
			for d in self.dumas[c]:
				self.annot_gere[c].append(self.ax[-1].annotate(d, xy=(self.B.range_dumas[c][d]['min'], 50*ay_), xytext=(-20,10),
					textcoords="offset points", bbox=dict(boxstyle="round", fc="w"), arrowprops=dict(arrowstyle="fancy")))

		for an_ in self.annot:
			an_.set_visible(False)
			an_.set_zorder(10)

		for c in self.B.range_dumas.keys():
			for an_ in self.annot_gere[c]:
				an_.set_visible(False)

		# legend pick event
		self.fig.canvas.mpl_connect('pick_event', self.onpick)
		leg_dumas.pick()

		self.fig.canvas.mpl_connect("button_press_event", self.click_)
		
		# ax[-1].text(105500, 49, 'IRS', fontdict={'fontsize':5})
		# buttonaxes = plt.axes([0.89, .95, 0.1, 0.025])
		# bnext = Button(buttonaxes, 'Export')
		# bnext.on_clicked(ExportToParallelCoordGraph)

		plt.show()

	def draw_vspans(self, ax_, range_dumas):
		from numpy import random
		# global range_dumas
		# random.seed(1234)
		ymax = 50
		for c in range_dumas.keys():
			if c == self.duma_Genome:
				ymax = 0.7
			elif c == self.duma_Repeat:
				ymax = 0.85
			else:		
				ymax = 1
			for d in range_dumas[c].keys():
				# r, g, b = random.uniform(0, 1, 3)
				self.spans[c+str(d)] = self.spans.get(c+str(d), [])

				self.spans[c+str(d)].append(ax_.axvspan(range_dumas[c][d]['min'], range_dumas[c][d]['max'], 
					0, ymax, color=scols.get_color(), alpha=0.1))

		scols.cur = 0

	def update_xlim(self, val):
		if self.zslider.valleft > self.zslider.valright:
			return
		for x in self.ax:
			x.axes.set_xlim([self.zslider.valleft, self.zslider.valright])
		self.fig.canvas.draw_idle()	

	def onpick(self, event):
		patch = event.artist
		vis = True if patch.get_alpha()==0.3 else False
		label = patch.get_label()
		color = 'grey'

		# basepair 클릭시 일관성을 위해... (필터링과)
		# color랑 몇가지 변수만 잘 설정해서 filtering로 넘어가기
		if label == 'A' or label == 'G' or label == 'C' or label == 'T':
			nmajor = self.basepair.index( label )
			color = self.basecolors[nmajor]
			self.major_visiable[ nmajor ] = vis
			for sc in self.sct:
				for mj, mn in zip(sc[nmajor]['major'], sc[nmajor]['minor']):
					mj.set_visible( vis )
					mn.set_visible( vis )
		elif label == self.duma_ORF or label ==self.duma_Repeat or label ==self.duma_Genome:
			for d in self.B.range_dumas[label].keys():
				for sc in self.spans[label+str(d)]:
					sc.set_visible(vis)
			for an_ in self.annot_gere[label]:
				an_.set_visible(vis)
		else:
			if label == 'base' : self.bidx += 1
			else: self.vidx += 1
			if self.bidx == len(self.base_range): self.bidx = 0
			if self.vidx == len(self.variation_range): self.vidx = 0

			range_, pid, color = (str(self.base_range[self.bidx]), 0, 'black') if label == 'base' else (str(self.variation_range[self.vidx]), 1, 'magenta')
			
			self.leg_filter.get_legend().get_texts()[pid].set_text(label +'_'+ range_)

			for i in range(len(self.bpdiff)):
				for a in range(4):
					if not self.major_visiable[a]: continue
					cond1 = self.bpdiff[i]['Mj_seq_x'] == a			

					if len( self.bpdiff[i][ cond1.values ] ) > 0:
						cond2 = logical_and(self.bpdiff[i]['maf_x'] >= self.base_range[self.bidx], abs(self.bpdiff[i]['maf_y'] - self.bpdiff[i]['maf_x']) >=  self.variation_range[self.vidx])					
						vis_d = logical_and(cond1, cond2)

						from numpy import isin
						for scmj, scmn in zip(self.sct[i][a]['major'], self.sct[i][a]['minor']):
							mj_vis_arr = isin(scmj.get_offsets()[:,0], self.bpdiff[i][vis_d.values][self.duma_position].tolist() )
							mn_vis_arr = isin(scmn.get_offsets()[:,0], self.bpdiff[i][vis_d.values][self.duma_position].tolist() )

							mjcolors_ = scmj.get_facecolors()
							mncolors_ = scmn.get_facecolors()
							if len(mjcolors_) == len(mj_vis_arr):
								mjcolors_[:,-1][ mj_vis_arr ] = 1
								mjcolors_[:,-1][ logical_not(mj_vis_arr) ] = 0
							else:
								mjc = list(mjcolors_[0][:-1])
								# !!
								mjcolors_ = [ mjc + [1.] if mi else mjc+[0.]  for mi in mj_vis_arr ]

							if len(mncolors_) == len(mn_vis_arr):
								mncolors_[:,-1][ mn_vis_arr ] = 1
								mncolors_[:,-1][ logical_not(mn_vis_arr) ] = 0
							
							scmj.set_facecolors(mjcolors_)
							scmj.set_edgecolors(mjcolors_)
							scmn.set_facecolors(mncolors_)
							scmn.set_edgecolors(mncolors_)
		if vis:
			patch.set_alpha(1.0)
		else:
			patch.set_facecolor(color)
			patch.set_alpha(0.3)
		self.fig.canvas.draw_idle()

	def click_(self, event):
		if event.xdata is None:
			return
		pos = int(event.xdata)
		# print(pos)
		epos = 20.
		minpos = pos - epos
		maxpos = pos + epos
		for i in range(len(self.bpdiff)):
			self.annot[i].set_visible(False)
			# bpdiff 0에서 찾기 ->  일관성..
			# binary search pos
			ind_ = self.bpdiff[i][[self.duma_position]].values.ravel().searchsorted(pos, side='left')
			if ind_ >= len(self.bpdiff[i][[self.duma_position]].values):
				ind_ = ind_ - 1
				
			spos = self.bpdiff[i][[self.duma_position]].values[ ind_ ]
			ind_1 = ind_ - 1
			spos_1 = self.bpdiff[i][[self.duma_position]].values[ ind_1 ]
			# print(ind_, spos, spos_1)
			ind_, spos = (ind_, spos) if abs(pos-spos) < abs(pos-spos_1) else (ind_1,spos_1)
			# print(ind_, spos)

			# print(bpdiff[i].values[ind_])
			ge = self.bpdiff[i][[self.duma_Genome]].values[ind_][0]
			re = self.bpdiff[i][[self.duma_Repeat]].values[ind_][0]
			orf_ = self.bpdiff[i][[self.duma_ORF]].values[ind_][0]
			# print(ge, re, orf_)

			bp_ = self.bpdiff[i].iloc[ind_]
			if spos < minpos or maxpos < spos:
				pass
				# tagging

			else: # minpos <= spos and spos <= maxpos:
				for sc_ in self.sct[i]:
					for s_ in sc_['major']:
						if s_ == None:
							continue
						# xy data -> xy (position in scatter)로 변환하는 방법
						if s_.get_offsets().__contains__(spos):
							ind = where(s_.get_offsets() == spos)
							pos_ = s_.get_offsets()[ind[0]]
							pos_[0][1] = bp_['maf_y']
							# print(pos_)
							self.annot[i].xy = pos_[0]
							text = "mj/mn : {}/{}\nmaf :{}->{}\nGnSt : {}\nRpRg : {}\nORF : {}".format(self.basepair[int(bp_[['Mj_seq_y']].values[0])], 
								self.basepair[int(bp_[['Mn_seq_y']].values[0])].lower(), round(float(bp_['maf_x']), 2), round(float(bp_['maf_y']), 2), 
								ge,re,orf_)
							self.annot[i].set_text(text)
							self.annot[i].get_bbox_patch().set_facecolor(self.basecolors[self.basepair.index(s_.get_label())])
							self.annot[i].get_bbox_patch().set_alpha(0.5)
							self.annot[i].set_visible(True)
		self.fig.canvas.draw_idle()

	def get_uppery(self, y):
		y_ = [ i[-1] for i in y ]
		return y_

	def draw_minor_transition(self, ax, sc, d, a, b):
		# transition from i to b
		for i in range(4):
			if i == b:
				continue
			d_ = d[ d['Mn_seq_x'] == i ]
			if len(d_) == 0: continue

			y_ = [ linspace(p_.maf_x, p_.maf_y, self.seq_size) for p_ in d_[['maf_x', 'maf_y']].itertuples() ]
			x_ = array(y_)
			for j in range(len(d_)):
				x_[j].fill( d_[self.duma_position].values[j] )
			c = [ self.xlin for _ in range(len(x_))]
			
			sc[a]['minor'].append(ax.scatter(x_, y_, c=c, cmap=mcols.cmap(self.dcolor[i][b]), s=self.markersize))

			# draw major scatter
			y_ = self.get_uppery(y_)
			x_ = [ i[0] for i in x_]
			sc[a]['major'].append(ax.scatter(x_, y_, color=self.basecolors[a], label=self.basepair[a], s=self.markersize-1))

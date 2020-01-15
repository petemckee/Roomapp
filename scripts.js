
var Roomapp = {

	includeSecondSlide: true,
	invalidPupilNames: ['Marksheet Name', 'Group Name', 'Export Date'],

    init: function() {

		var me = this;
		this.cHead = document.querySelector('.js-head');

        document.querySelector('input').addEventListener('change', function() {
			me.parseExcel( this.files[0] );
			me.cHead.classList.add('active');
        });

		document.querySelector('.js-create').addEventListener('click', function(e) {
			me.createPpt(me.students);
		});

        this.$resultsArea = document.querySelector('.js-students');
		this.$className = document.querySelector('.js-className');

		this.handleSelections();

    },
    formatName: function(name) {
		var splitName = name.split(' ');
		// TODO -> Check for dupes of surname
		return splitName.reverse()[0] + ' ' + splitName[1].slice(0,1);
	},
	getPupilPremium: function(d) {
		var pupilPremium = d["Pupil Premium Indicator"];
		return pupilPremium === 'Y' ? 'PP ' : '';
	}
	, getSen: function(d) {
		var senStatus = d["SEN Status"];
		return senStatus !== '' ? senStatus + ' ' : '';
	},
	getTargetGradeColumnName: function(colNames) {
        // TODO -> Run this on first row to get col index 
		var col = colNames.filter(function(x) {
			return (x.startsWith('Year') && x.indexOf('Target') !== -1)
		});
		return (col.length === 1) ? col[0] : null;
	},
	getCurrentGradeColumnName: function(colNames) {
        // TODO -> Run this on first row to get col index 
		var col = colNames.filter(function(x) {
			return (x.startsWith('Year') && x.indexOf('Band') !== -1)
		});
		return (col.length === 1) ? col[0] : null;
	},
	getColumnValue: function(val) {
		return (val !== undefined) ? val : '';
	},
	parseAndDisplayJson: function(data) {
	
		var me = this;
		// -> TODO get just the data we need...
		// -> TODO use mustache template

		var firstRow = data[0];
		var colNames = Object.getOwnPropertyNames(firstRow);
		var targetGradeColumnName = me.getTargetGradeColumnName(colNames);
		var currentGradeColumnName = me.getCurrentGradeColumnName(colNames);
				
		var students = data.map(function(d) {
			var name = me.formatName(d['Surname Forename']);
			// TODO -> Check if array worked and only has 2 elements
			// TODO -> Check surnames to see if first letter is required...
			/* 	var studentDetails = String.Format("{0}T: {1} {2}{0}{3}{4}", Environment.NewLine, student.TargetGrade, student.CurrentGrade, pp, sen*/
							
			return {
				name: name,
				pupilPremium: me.getPupilPremium(d),
				senStatus: me.getSen(d),
				targetGrade: me.getColumnValue(d[targetGradeColumnName]),
				currentGrade: me.getColumnValue(d[currentGradeColumnName]),
				//include: true
				include: me.looksLikeValidPupil(d['Surname Forename'])
			}
		});
		
		// TODO combine into map above...
		students = students.map(function(s) {
			return {
				name: me.getName(s, students),
				firstLine: me.firstLine(s),
				secondLine: me.secondLine(s),
				include: true
			}
		})
		
		me.students = students;
		
		var $html = '<table><thead><tr><th class="name">Name</th><th class="details">Details</th><th class="include">Include<span class="js-includeNo includeNo">[ '+students.length+' ]</span></th></tr></thead><tbody>'
		for (var i = 0; i < students.length; i++) {
			$html += '<tr><td>' + students[i].name + '</td><td>' + students[i].firstLine + ', ' + students[i].secondLine + '</td>' +
			'<td class="chk"><span class="chkBox" data-item-index="' + i + '"></span><input type="checkbox" checked="checked" name="include" data-item-index="'+i+'" /></td></tr>';
		}
		$html += '</tbody></table>';
		
		var newdiv = document.createElement('div');
        newdiv.innerHTML = $html;
        
        this.$className.value = data[0]['Class'];
		this.$resultsArea.innerHTML = newdiv.innerHTML;

		this.$includeNo = document.querySelector('.js-includeNo');
    },
	
	getName: function(student, data) {

		console.log(student, data);
		return student.name;

	}, 

	looksLikeValidPupil: function(name) {
	

		return this.invalidPupilNames.some(function(v) { return name.indexOf(v) === -1; });

		//substrings.some(function(v) { return str.indexOf(v) >= 0; })

	},

	handleSelections: function() { 
	
		var me = this;
		document.body.addEventListener('click', function(e) {
            
			if (e.target.tagName === 'TD') {
				var row = e.target.parentElement;
                var chk = row.querySelector('input[type=checkbox]');
                
                if (chk.checked) {
                    row.classList.add('disabled');
                    chk.checked = false;
                } else {
                    row.classList.remove('disabled');
                    chk.checked = true;    
                }				
				me.students[chk.getAttribute('data-item-index')]['include'] = chk.checked;
			}


			me.$includeNo.innerText = '[ '+  me.$resultsArea.querySelectorAll('input[type=checkbox]:checked').length +' ] ';

		});
	
	},	
	createPpt: function(data) {
	
		data = data.filter(function(i) { return i.include === true });
				
		var pptx = new PptxGenJS();
		var slide = pptx.addNewSlide();

		slide.addText(this.$className.value, { x: 0, y: 0, h: 0.5, w: 1.5, fontSize: 12, color: '000000' });

		if (this.includeSecondSlide) {
			var slide2 = pptx.addNewSlide();
			slide2.addText(this.$className.value, { x: 0, y: 0, h: 0.5, w: 1.5, fontSize: 12, color: '000000' });
		}

        // pptxGen measurements are in inches
        this.boxWidth = 1.25;
        var positions = this.getSeatPositions();    

		for (var i = 0; i < data.length; i++) {
			
            var student = data[i];
            var position = positions[i];

			slide.addText([
				{text: student.name, options: { fontSize: 10, color: '000000', breakLine: true }},
				{text: student.firstLine, options: { fontSize: 8, color: '000000', breakLine: true }},
				{text: student.secondLine, options: { fontSize: 8, color: '000000'}}
				],
				{ x: position.x, y: position.y, h: 0.5, w: this.boxWidth, fontSize: 10, color: '000000', line: '303030' }
			);

			slide2.addText([{text: student.name, options: { fontSize: 10, color: '000000', breakLine: true }}],
				{ x: position.x, y: position.y, h: 0.5, w: this.boxWidth, fontSize: 10, color: '000000', line: '303030' }
			);
		}

		pptx.save('demo');
    },

	firstLine: function(student) {
		return 'T: ' + student.targetGrade + ' ' + student.currentGrade
	},
	secondLine: function(student) {
		return ((student.pupilPremium) ? student.pupilPremium + ' ' : '') + student.senStatus
	},

    getSeatPositions: function() {

        // TODO - Eventually get positions from template/pptx/?
		var colLeftMargin = 0;
        var col = 0;
        var xCoor = 0;
        var yCoor = 5.5;
        var width = this.boxWidth;
        var positions = [];

        for (var i = 0; i < 30; i++) {
			
            if (i % 8 === 0) {
                xCoor = colLeftMargin;
				col = 0;	
                yCoor = yCoor - 1;
            } else {
                col++;
                xCoor = parseFloat((col * width + colLeftMargin).toFixed(2));
            }

            positions.push({ x: xCoor, y: yCoor });
        }

        return positions;
    },

    parseExcel: function(file) {

        var me = this;
        var reader = new FileReader();
    
        reader.onload = function(e) {
          var data = e.target.result;
          var workbook = XLSX.read(data, {
            type: 'binary'
          });
    
          workbook.SheetNames.forEach(function(sheetName) {
            var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            me.parseAndDisplayJson(XL_row_object);
          });
    
        };
    
        reader.onerror = function(ex) {
          console.log(ex);
        };
    
        reader.readAsBinaryString(file);
      }
}

Roomapp.init();
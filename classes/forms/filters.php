<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Version metadata for the report_moduleassignment plugin.
 *
 * @package   report_moduleassignment
 * @copyright 2024, Universindad Ciudadana de Nuevo leon {@link http://www.ucnl.edu.mx/}
 * @license   http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 * @author    Adrian Francisco Lozada Reboceño <adrian.lozada@ucnl.edu.mx>
 */

namespace report_moduleassignment\forms;

defined('MOODLE_INTERNAL') || die();
require_once($CFG->libdir . '/formslib.php');

class filters extends \moodleform {
    public function definition() {
        global $USER, $PAGE, $DB, $CFG;
        $mform =& $this->_form;

        $academicdegree = get_config('report_ucnl', 'academicdegree');         
        $category = $this->get_categories();
        switch ($academicdegree) {
            case '2':
                # code... 
                $attributes = ["onchange" => "javascript:get_program(this.value,false)", ];              
                $mform->addElement('select',  'period',  get_string('period', 'report_moduleassignment'),  $category, $attributes);   // periodo
                $attributes = ["onchange" => "javascript:get_semester(this.value,false)", ];
                $mform->addElement('select',  'program',  get_string('program','report_moduleassignment'),'', $attributes); // programa */
                break;
            case '3':
                # code...
                $attributes = ["onchange" => "javascript:get_program(this.value,false)", ];              
                $mform->addElement('select',  'period',  get_string('period', 'report_moduleassignment'),  $category, $attributes);   // periodo
                $attributes = ["onchange" => "javascript:get_semester(this.value,false)", ];
                $mform->addElement('select',  'program',  get_string('program','report_moduleassignment'),'', $attributes); // programa */
                break;
            default:
                # code...
                $attributes = ["onchange" => "javascript:get_semester(this.value,false)", ];              
                $mform->addElement('select',  'period',  get_string('period', 'report_moduleassignment'),  $category, $attributes);   // periodo              
                break;
        }
        $attributes = ["onchange" => "javascript:get_course(this.value,false)", ];
        $mform->addElement('select',  'semester',  get_string('semester','report_moduleassignment'),'', $attributes); // semestre */    
        $attributes = ["onchange" => "javascript:get_teacher(this.value,false)", ];
        $mform->addElement('select',  'course',  get_string('course'),'' , $attributes);

        $mform->addElement('submit','save_filter_form',get_string('filter'));
    }

    public function get_categories() {
        global $DB;
        $categories = $DB->get_records('course_categories', array('parent' => 0));
        $data[0] = 'Todos';
        foreach ($categories as $category) {
            $data[$category->id] = $category->name;
        }
        return $data;
    }
    public function get_subcategories($category) {
        global $DB;
        $subcategories = $DB->get_records('course_categories', array('parent' => $category));
        $subcategories = array('0' => get_string('all')) + $subcategories;
        return $subcategories;
    }
    public function get_data(){
        $data = parent::get_data();
    
        if(!empty($data)){
            $mform =& $this->_form;
            if(!empty($mform->_submitValues['period'])) {
                $data->period = $mform->_submitValues['period'];
            }
            if(!empty($mform->_submitValues['program'])) {
                $data->program = $mform->_submitValues['program'];
            }
            if(!empty($mform->_submitValues['semester'])) {
                $data->semester = $mform->_submitValues['semester'];
            }
            if(!empty($mform->_submitValues['course'])) {
                $data->course = $mform->_submitValues['course'];
            }
            return $data;
        }
    }
}
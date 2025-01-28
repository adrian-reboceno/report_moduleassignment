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


define('REPORT_MODULEASSIGNMENT_ACTION_LOAD_FILTER', 'loadfilter');
define('REPORT_MODULEASSIGNMENT_ACTION_QUICK_FILTER', 'quickfilter');
define('FIXED_NUM_COLS', 6);

require_once('../../config.php');
require($CFG->libdir . '/phpspreadsheet/vendor/autoload.php');
    
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use core\notification;

use report_moduleassignment\forms\filters as form;


function report_moduleassignment_filter_form_action($filterid = null, $data = [], $quickfilter = false) {
    global $CFG, $PAGE, $USER;
    
    $customdata = [           
        'quickfilter' => $quickfilter,
        'userid'      => (int)$USER->id // For the hidden userid field.
    ];
    $customdata = array_merge($customdata, $data);
    $action     = $quickfilter ? REPORT_MODULEASSIGNMENT_ACTION_QUICK_FILTER : REPORT_MODULEASSIGNMENT_ACTION_LOAD_FILTER;
    $filterform = new form($PAGE->url->out(false) . '?action=' . $action, $customdata);

    return $filterform;
}

function report_moduleassignment_get_lastaccess($filter){
    global $DB, $USER;
    $plugin = 'report_moduleassignment';
    $columns = [
        get_string('id',  $plugin),
        get_string('firstname'),
        get_string('lastname'),
        get_string('email'),
        get_string('courseid',  $plugin),        
        get_string('fullname'),
        get_string('shortname'),
        get_string('category'),
        get_string('startdate'),
        get_string('enddate'),
        get_string('namesection', $plugin),
        get_string('typemodule', $plugin),
        get_string('moduloname', $plugin),
        get_string('start', $plugin),
        get_string('end', $plugin),
    ];
    $datarows['thead'] = $columns;
    if(!empty($filter->period)){       
        $period = (int)$filter->period;
        $where= "(cats.path LIKE '%/{$period}%' OR cats.path LIKE '%/{$period}' )";
    }
    if(!empty($filter->program)){                   
        $programa= (int)$filter->program;
        $where= "(cats.path LIKE '%/{$programa}%' OR cats.path LIKE '%/{$programa}' )";
    }
    if(!empty($filter->semester)){        
        $semester = (int)$filter->semester;
        $where= "(cats.path LIKE '%/{$semester}%' OR cats.path LIKE '%/{$semester}' )";
    }
    if(!empty($filter->course)){
        $course = (int)$filter->course;
        $where = "c.id = {$course}";
    }    
    if (!empty($where)) {
        $where = ' AND '. $where;
    } else {
        $where = '';
    }
    
    $sql ="(SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, q.name as activityname, 
                q.timeopen AS startactivity,
                q.timeclose AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_quiz AS q ON q.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'quiz'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, q.name, q.timeopen, q.timeclose)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, f.name as activityname, 
                f.duedate AS startactivity,
                f.cutoffdate AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_forum AS f ON f.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'forum'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, f.name, f.duedate, f.cutoffdate)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, ass.name as activityname, 
                ass.duedate AS startactivity,
                ass.cutoffdate AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_assign AS ass ON ass.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'assign'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, ass.name, ass.duedate, ass.cutoffdate)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, chat.name as activityname, 
                chat.chattime AS startactivity,
                NULL AS  endactivity               
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_chat AS chat ON chat.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'chat'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, chat.name, chat.chattime)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, choice.name as activityname, 
                choice.timeopen AS startactivity,
                choice.timeclose AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_choice AS choice ON choice.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'choice'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, choice.name, choice.timeopen, choice.timeclose)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, book.name as activityname, 
                NULL AS startactivity,
                NULL AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_book AS book ON book.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'book'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, book.name)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, dat.name as activityname, 
                dat.timeavailablefrom AS startactivity,
                dat.timeavailableto AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_data AS dat ON dat.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'data'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, dat.name, dat.timeavailablefrom, dat.timeavailableto)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, h5pac.name as activityname, 
                NULL AS startactivity,
                NULL AS  endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_h5pactivity AS h5pac ON h5pac.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'h5pactivity'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, h5pac.name)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, imscp.name as activityname, 
                NULL AS startactivity,
                NULL AS  endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_imscp AS imscp ON imscp.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'imscp'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, imscp.name)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, scorm.name as activityname,                 
                scorm.timeopen AS startactivity,
                scorm.timeclose AS endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_scorm AS scorm ON scorm.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'scorm'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, scorm.name,  scorm.timeopen,  scorm.timeclose)
        UNION
            (SELECT
                u.id AS id, u.firstname, u.lastname, u.email,
                c.id AS idcurso, c.fullname, c.shortname, cats.name AS namecategory, c.startdate, c.enddate,
                cs.section, cs.name AS 'namesection', m.name AS 'typemodule', cm.id, cm.instance, cm.section, survey.name as activityname, 
                NULL AS startactivity,
                NULL AS  endactivity
            FROM mdl_course AS c
            LEFT JOIN mdl_context AS ctx ON c.id = ctx.instanceid
            JOIN mdl_role_assignments AS lra ON lra.contextid = ctx.id
            JOIN mdl_role_assignments AS tra ON tra.contextid = ctx.id
            JOIN mdl_user AS u ON lra.userid = u.id
            JOIN mdl_course_categories AS cats ON c.category = cats.id
            JOIN mdl_course_sections AS cs ON cs.course = c.id AND cs.section <= 20 AND cs.section >= 0
            LEFT JOIN mdl_course_modules AS cm ON cm.course = c.id AND cm.section = cs.id
            JOIN mdl_modules AS m ON m.id = cm.module AND m.name NOT IN ('label', 'url', 'page', 'resource')
            LEFT JOIN mdl_survey AS survey ON survey.id = cm.instance
            WHERE
                lra.roleid IN (3) 
                AND m.name LIKE 'survey'
                {$where}
            GROUP BY u.id, u.firstname, u.lastname, u.email, c.id, c.fullname, c.shortname, cats.name, c.startdate, c.enddate, cs.section, cs.name, m.name, cm.id, cm.instance, cm.section, survey.name)
        ORDER BY c.id, cs.section";

    $data = $DB->get_records_sql($sql, $params);

    $day  =  date("Y-m-d H:i:s");
    $daytotal=0;
    $day = new DateTime (substr($day,0,10));
    $daysinactivitystudents = get_config('report_ucnl', 'daysinactivitystudents');
    foreach ($data as $key => $value) {
        $daytotal=0;
        $row['id'] = $value->id;
        $row['firstname'] = $value->firstname;
        $row['lastname'] = $value->lastname;
        $row['email'] = $value->email;
        $row['idcurso'] = $value->idcurso;
        $row['fullname'] = $value->fullname;
        $row['shortname'] = $value->shortname;
        $row['namecategory'] = $value->namecategory;
        $row['startdate'] =  date('Y-m-d H:i:s', $value->startdate);
        $row['enddate'] =  date('Y-m-d H:i:s', $value->enddate);
        $row['section'] = $value->section;
        $row['namesection'] = $value->namesection;
        $row['typemodule'] = $value->typemodule;
        $typemodule = $value->typemodule;       
        $mparams['instance'] = (int)$value->instance;
        $mparams['course'] = (int)$value->idcurso;
        $datamodule = $DB->get_record_sql($sqlmodule, $mparams);
        $row['moduloname'] = null;
        $row['qualified'] = null;
        if($typemodule=='forum'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin);
            if(empty($value->startactivity) ||  ($value->endactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;
            }
            else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='quiz'){
            $row['moduloname'] = $value->activityname;         
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin);
            if(empty($value->startactivity) ||  ($value->startactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;
                $row['qualified'] = '';
            } else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='assign'){
            $row['moduloname'] = $value->activityname;         
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin); 
            if(empty($value->startactivity) ||  ($value->endactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;
            }
            else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='chat'){
            $row['moduloname'] = $value->activityname;         
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = '';
            if(empty($value->startactivity) ||  ($value->startactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;               
            }else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='choice'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin);
            if(empty($value->startactivity) ||  ($value->endactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;                
            }else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='book'){
            $row['moduloname'] = $value->activityname;
            $row['start'] ='';
            $row['end'] = '';          
            $row['qualified'] = '';          
        }elseif($typemodule=='data'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin);
            if(empty($value->startactivity) ||  ($value->endactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;              
            }else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='h5pactivity'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = '';
            $row['end'] = '';          
            $row['qualified'] = '';
        }elseif($typemodule=='imscp'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = '';
            $row['end'] = '';          
            $row['qualified'] = '';
        }elseif($typemodule=='scorm'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = $value->startactivity ? date('Y-m-d H:i:s', $value->startactivity ) : get_string('indefinite', $plugin);
            $row['end'] = $value->endactivity ? date('Y-m-d H:i:s', $value->endactivity) : get_string('indefinite', $plugin);
            if(empty($value->startactivity) ||  ($value->endactivity < $value->startdate)){
                $qualified = get_config('report_ucnl', 'notqualified');
                $row['qualified'] = $qualified;                
            }else{
                $row['qualified'] = '';
            }
        }elseif($typemodule=='survey'){
            $row['moduloname'] = $value->activityname;
            $row['start'] = '';
            $row['end'] = '';          
            $row['qualified'] = '';
        }
       /*  var_dump($datamodule); */       
        $datarows['rows'][] = $row;
    }

    return $datarows;
}

function report_excel($data){
    global $CFG, $DB;
    $now = time();
    $spread = new Spreadsheet();

    $spread->getActiveSheet()->getStyle('A1:O1')->getFont()->setBold(true);
    foreach (range('A','O') as $col) {
        $spread->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
    }
    foreach($data['thead'] as $key => $value){
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow($key+1, 1, $value);
    }
    foreach($data['rows'] as $key => $value){
        if($value['qualified']){
            $spread->getActiveSheet()->getStyle('N'. $key+2)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB( strtoupper(str_replace('#','',$value['qualified'])));
            $spread->getActiveSheet()->getStyle('O'. $key+2)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB( strtoupper(str_replace('#','',$value['qualified'])));
        }       
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(1, $key+2, $value['id']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(2, $key+2, $value['firstname']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(3, $key+2, $value['lastname']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(4, $key+2, $value['email']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(5, $key+2, $value['idcurso']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(6, $key+2, $value['fullname']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(7, $key+2, $value['shortname']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(8, $key+2, $value['namecategory']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(9, $key+2, $value['startdate']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(10, $key+2, $value['enddate']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(11, $key+2, $value['namesection']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(12, $key+2, $value['typemodule']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(13, $key+2, $value['moduloname']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(14, $key+2, $value['start']);
        $spread->setActiveSheetIndex(0)->setCellValueByColumnAndRow(15, $key+2, $value['end']);
    }
    $spread->getActiveSheet()->setTitle(get_string('pluginname', 'report_moduleassignment'));
    $writer = new Xlsx($spread);
    $filename = 'report_'.$now.'.xlsx'; 
    $filename = 'export/'.$filename;  
    $writer->save($filename);
    return $filename;
}



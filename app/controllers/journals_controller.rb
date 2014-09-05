# Redmine - project management software
# Copyright (C) 2006-2014  Jean-Philippe Lang
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

require 'rubygems'
require 'spreadsheet'

class JournalsController < ApplicationController
  before_filter :find_journal, :only => [:edit, :diff]
  before_filter :find_issue, :only => [:new]
  before_filter :find_optional_project, :only => [:index]
  before_filter :authorize, :only => [:new, :edit, :diff]
  accept_rss_auth :index
  menu_item :issues

  helper :issues
  include IssuesHelper
  helper :custom_fields
  include CustomFieldsHelper
  helper :queries
  include QueriesHelper
  helper :sort
  include SortHelper

  def index
    retrieve_query
    sort_init 'id', 'desc'
    sort_update(@query.sortable_columns)

    if params[:format]=='xls'       
      if !params[:startTime].nil? && !params[:endTime].nil? && !params[:startTime].empty? && !params[:endTime].empty?
        logger.info("********##*********")
        logger.info(params[:startTime])
        logger.info(params[:endTime])
        logger.info("********************")
        startTime = Date::parse(params[:startTime])
        endTime = Date::parse(params[:endTime])+1.day
        @journals = Journal.searchJournals(startTime,endTime)
      else
        @journals = nil
      end



      Spreadsheet.client_encoding = "UTF-8"
      book = Spreadsheet::Workbook.new
      sheet1 = book.create_worksheet :name => "task list"
      default_format = Spreadsheet::Format::new(:weight => :bold, :size => 10,
                                                :align => :merge, :color => "black", :border => :thin,
                                                :border_color => "black", :pattern => 1,
                                                :pattern_fg_color => "yellow", :text_wrap => 1)

      data_format = Spreadsheet::Format::new( :size => 10,:color => "black", :border => :thin,
                                                :border_color => "black", :pattern => 1,
                                                :pattern_fg_color => "white", :text_wrap => 1)
      
      header_row = sheet1.row(0)

      8.times do |i|
        sheet1.column(i).width=12
        header_row.set_format(i,default_format)
      end
      sheet1.column(1).width=15
      sheet1.column(2).width=15
      sheet1.column(3).width=48
      sheet1.column(5).width=28
      sheet1.column(6).width=28

      header_row[0] = l(:label_date)
      header_row[1] = l(:revisor)
      header_row[2] = l(:renwubianhao)
      header_row[3] = l(:gongcheng)+"/"+l(:taskname)
      header_row[4] = l(:entry)
      header_row[5] = l(:change_before)
      header_row[6] = l(:change_after)
      header_row[7] = l(:beizhu)

      j = 0
      k = 0
      old_user = ""
      old_data = ""
      

      issueTaskNoField = CustomField.find_by_name_and_type(l(:renwubianhao),"IssueCustomField")

      @journals.each do |journal|

          journal.visible_details.each do |detail|
            old_value = ""
            value = ""
            label = ""

            j = j+1
            data_row = sheet1.row(j)
            data_row[0] = journal.created_on.strftime("%Y-%m-%d")
            data_row[1] = journal.user.firstname + journal.user.lastname
            logger.info("*****"+issueTaskNoField.id.to_s+"******"+journal.journalized_id.to_s+"")
                       
           
            if journal.journalized_type=='Issue'

              if detail.property=='attr' && detail.prop_key=='destroy_issue'
                data_row[2] = detail.old_value
                data_row[3] = detail.value
              else
                custom_value = CustomValue.find_by_customized_id_and_custom_field_id(journal.journalized_id,issueTaskNoField.id)
                if custom_value.nil?
                  j = j-1
                  next
                end
                # renwubianhao
                data_row[2] = custom_value.value
                data_row[3] = journal.issue.subject
              end
              
            else
              
              data_row[2] = detail.old_value 
              data_row[3] = detail.value

            end

            case detail.property
            when "cf"
              field = detail.custom_field
      
              label = field.name
              logger.info("******"+field.name+"*******")
              if detail.old_value
                old_value = format_value(detail.old_value, field)
              end
              if detail.value
                value = format_value(detail.value, field)
              end

            when "attr"
              field = detail.prop_key.to_s.gsub(/\_id$/, "")
              label = l(("field_" + field).to_sym)
              case detail.prop_key
              when 'due_date', 'start_date'
                value = format_date(detail.value.to_date) if detail.value
                old_value = format_date(detail.old_value.to_date) if detail.old_value

              when 'project_id', 'status_id', 'tracker_id', 'assigned_to_id',
                    'priority_id', 'category_id', 'fixed_version_id'
                value = find_name_by_reflection(field, detail.value)
                old_value = find_name_by_reflection(field, detail.old_value)

              when 'estimated_hours'
                value = "%0.02f" % detail.value.to_f unless detail.value.blank?
                old_value = "%0.02f" % detail.old_value.to_f unless detail.old_value.blank?

              when 'parent_id'
                label = l(:field_parent_issue)
                value = "##{detail.value}" unless detail.value.blank?
                old_value = "##{detail.old_value}" unless detail.old_value.blank?

              when 'is_private'
                value = l(detail.value == "0" ? :general_text_No : :general_text_Yes) unless detail.value.blank?
                old_value = l(detail.old_value == "0" ? :general_text_No : :general_text_Yes) unless detail.old_value.blank?

              when 'subject','done_ratio'
                old_value = detail.old_value
                value = detail.value

              when 'create_project','destroy_project','create_issue','destroy_issue'
                old_value = ""
                value = ""
                data_row[4] = label
                data_row[5] = old_value
                data_row[6] = value
                next
              end

            when 'attachment'
              label = l(:label_attachment)

            when 'relation'
              if detail.value && !detail.old_value
                rel_issue = Issue.visible.find_by_id(detail.value)
                value = rel_issue.nil? ? "#{l(:label_issue)} ##{detail.value}" :
                          (no_html ? rel_issue : link_to_issue(rel_issue, :only_path => options[:only_path]))
              elsif detail.old_value && !detail.value
                rel_issue = Issue.visible.find_by_id(detail.old_value)
                old_value = rel_issue.nil? ? "#{l(:label_issue)} ##{detail.old_value}" :
                                  (no_html ? rel_issue : link_to_issue(rel_issue, :only_path => options[:only_path]))
              end
              relation_type = IssueRelation::TYPES[detail.prop_key]
              label = l(relation_type[:name]) if relation_type

            end

            label ||= detail.prop_key
            value ||= detail.value
            old_value ||= detail.old_value

            data_row[4] = label
            data_row[5] = old_value
            data_row[6] = value
          end
       
      end

      book.write 'files/changes_list.xls'
      send_file('files/changes_list.xls')
    else
      logger.info()
      if @query.valid?
        @journals = @query.journals(:order => "#{Journal.table_name}.created_on DESC",
                                    :limit => 25)
      end
      @title = (@project ? @project.name : Setting.app_title) + ": " + (@query.new_record? ? l(:label_changes_details) : @query.name)
      render :layout => false, :content_type => 'application/atom+xml'
    end
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  def diff
    @issue = @journal.issue
    if params[:detail_id].present?
      @detail = @journal.details.find_by_id(params[:detail_id])
    else
      @detail = @journal.details.detect {|d| d.prop_key == 'description'}
    end
    (render_404; return false) unless @issue && @detail
    @diff = Redmine::Helpers::Diff.new(@detail.value, @detail.old_value)
  end

  def new
    @journal = Journal.visible.find(params[:journal_id]) if params[:journal_id]
    if @journal
      user = @journal.user
      text = @journal.notes
    else
      user = @issue.author
      text = @issue.description
    end
    # Replaces pre blocks with [...]
    text = text.to_s.strip.gsub(%r{<pre>(.*?)</pre>}m, '[...]')
    @content = "#{ll(Setting.default_language, :text_user_wrote, user)}\n> "
    @content << text.gsub(/(\r?\n|\r\n?)/, "\n> ") + "\n\n"
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  def edit
    (render_403; return false) unless @journal.editable_by?(User.current)
    if request.post?
      @journal.update_attributes(:notes => params[:notes]) if params[:notes]
      @journal.destroy if @journal.details.empty? && @journal.notes.blank?
      call_hook(:controller_journals_edit_post, { :journal => @journal, :params => params})
      respond_to do |format|
        format.html { redirect_to issue_path(@journal.journalized) }
        format.js { render :action => 'update' }
      end
    else
      respond_to do |format|
        format.html {
          # TODO: implement non-JS journal update
          render :nothing => true
        }
        format.js
      end
    end
  end

  private

  def find_journal
    @journal = Journal.visible.find(params[:id])
    @project = @journal.journalized.project
  rescue ActiveRecord::RecordNotFound
    render_404
  end

end

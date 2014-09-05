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


class IssuesController < ApplicationController
  menu_item :new_issue, :only => [:new, :create]
  default_search_scope :issues

  before_filter :find_issue, :only => [:show, :edit, :update]
  before_filter :find_issues, :only => [:bulk_edit, :bulk_update, :destroy]
  before_filter :find_project, :only => [:new, :create, :update_form]
  before_filter :authorize, :except => [:index]
  before_filter :find_optional_project, :only => [:index]
  before_filter :check_for_default_issue_status, :only => [:new, :create]
  before_filter :build_new_issue_from_params, :only => [:new, :create, :update_form]
  accept_rss_auth :index, :show
  accept_api_auth :index, :show, :create, :update, :destroy

  rescue_from Query::StatementInvalid, :with => :query_statement_invalid

  helper :journals
  helper :projects
  include ProjectsHelper
  helper :custom_fields
  include CustomFieldsHelper
  helper :issue_relations
  include IssueRelationsHelper
  helper :watchers
  include WatchersHelper
  helper :attachments
  include AttachmentsHelper
  helper :queries
  include QueriesHelper
  helper :repositories
  include RepositoriesHelper
  helper :sort
  include SortHelper
  include IssuesHelper
  helper :timelog
  include Redmine::Export::PDF

  def index
    retrieve_query

    sort_init(@query.sort_criteria.empty? ? [['id', 'desc']] : @query.sort_criteria)

    journals_hash = {}
    
    if @query.name == l(:renwudabiao)
    
      custom_field = CustomField.find_by_name_and_type(l(:suoyin),'IssueCustomField')

      journals = Journal.searchJournals()
      journals.each do |journal|
      	if journal.journalized_type=='Issue'
	  journals_hash[journal.journalized_id] = journal
	end
      end

      if params[:sort].nil?
      	params[:sort] = "cf_"+custom_field.id.to_s
	
      elsif params[:sort]!=("cf_"+custom_field.id.to_s)

      	params[:sort] = "cf_"+custom_field.id.to_s+"," + params[:sort] 
      end
      logger.info(params[:sort])
    end


    sort_update(@query.sortable_columns)
    @query.sort_criteria = sort_criteria.to_a

    


    if @query.valid?
      case params[:format]
      when 'csv', 'pdf','xls'
        @limit = Setting.issues_export_limit.to_i
        if params[:columns] == 'all'
          @query.column_names = @query.available_inline_columns.map(&:name)
        end
      when 'atom'
        @limit = Setting.feeds_limit.to_i
      when 'xml', 'json'
        @offset, @limit = api_offset_and_limit
        @query.column_names = %w(author)
      else
        @limit = per_page_option
      end

      @issue_count = @query.issue_count
      @issue_pages = Paginator.new @issue_count, @limit, params['page']
      @offset ||= @issue_pages.offset
      if @query.name == l(:renwudabiao)

      	@issues = @query.issues1(:include => [:assigned_to, :tracker, :priority, :category, :fixed_version],
                              :order => sort_clause)
      else
      	@issues = @query.issues(:include => [:assigned_to, :tracker, :priority, :category, :fixed_version],
                              :order => sort_clause,
			      :offset => @offset,
			      :limit => @limit)
	
      end
                              
      @issue_count_by_group = @query.issue_count_by_group


      curGrp = ""
      preGrp = ""
      @i = 0

      if !@issue_count_by_group.nil?

        @issues.each do |issue|
          curGrp = @query.group_by_column.value(issue)
         
          if @i==0
            preGrp = curGrp
          end

          if curGrp != preGrp 
            break
          end
          @i = @i+1
        end
      end

      if params[:format]=='xls'
     
        Spreadsheet.client_encoding = "UTF-8"
        book = Spreadsheet::Workbook.new
        sheet1 = book.create_worksheet :name => "task list"
        default_format = Spreadsheet::Format::new(:weight => :bold, :size => 11,
                                                  :align => :merge, :color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_33", :text_wrap => 1)

        data_format = Spreadsheet::Format::new( :size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_34", :text_wrap => 1)


        data_format1 = Spreadsheet::Format::new( :vertical_align => :middle ,:size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_34", :text_wrap => 1)

        data_format2 = Spreadsheet::Format::new( :horizontal_align => :center ,:vertical_align => :middle ,:size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_34", :text_wrap => 1)
        

        data_format3 = Spreadsheet::Format::new( :size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_14", :text_wrap => 1)
         
        data_format4 = Spreadsheet::Format::new( :size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_22", :text_wrap => 1)                                          

        data_format5 = Spreadsheet::Format::new( :size => 10,:color => "black", :border => :thin,
                                                  :border_color => "black", :pattern => 1,
                                                  :pattern_fg_color => "xls_color_5", :text_wrap => 1)                                          
        header_row = sheet1.row(0)        
        
        15.times do |i|
          sheet1.column(i).width=9
          header_row.set_format(i,default_format)
        end
	sheet1.column(0).width=4
        sheet1.column(1).width=20
        sheet1.column(2).width=6
	sheet1.column(3).width=4
        sheet1.column(4).width=6
        sheet1.column(6).width=27
        sheet1.column(7).width= 6
        sheet1.column(8).width= 5
        sheet1.column(9).width= 6
        sheet1.column(13).width=6
        sheet1.column(14).width=18

	


        header_row[0] = l(:xuhao)
        header_row[1] = l(:gongcheng)
        header_row[2] = l(:qiantouren)
        header_row[3] = l(:duban)
        header_row[4] = l(:dubanDeadline)
        header_row[5] = l(:renwubianhao)
        header_row[6] = l(:renwu)
        header_row[7] = l(:fuzeren)
        header_row[8] = l(:pingtai)
        header_row[9] = l(:pingtaifenlei)      
        header_row[10] = l(:yewubumen)
        header_row[11] = l(:shishituandui)
        header_row[12] = l(:renwujieduan)
        header_row[13] = l(:shijiandian)
        header_row[14] = l(:beizhu)
        
        j=0
        flag = 0
        temp_id = ""

   
        previous_group = ""

        @issues.each do |issue|

          j = j+1
          if @query.grouped? && (group = @query.group_by_column.value(issue)) != previous_group
            if @query.name==l(:renwudabiao)
              group_cnt = @issue_count_by_group[group] 
            
              if flag==0
                group_cnt = @i
                flag =1
              end
            end
            previous_group = group 
          end
          
      
          if @query.name==l(:renwudabiao)
            n=0
            columns = @query.inline_columns
         
            if temp_id == cus_column_content(columns[1], issue)
                 n=3
            end
            temp_id = cus_column_content(columns[1], issue)

            
            len = @query.inline_columns.size
            data_row = sheet1.row(j)
            15.times do |i|
              data_row.set_format(i,data_format)
            end
             
	      

            for i in 0..len-1 
              if @query.grouped?
                if n==0 && i <5      
		  if i==0
		    data_row.set_format(i,data_format2) 
    		  else
		    data_row.set_format(i,data_format1) 
    		  end
                           
                  sheet1.merge_cells(j,i,j+group_cnt-1,i)
                 
                  data_row[i] = cus_column_content(columns[i], issue)
                  

                elsif i >=5
		  if i==12 && cus_column_content(columns[i],issue)==l(:status_finished)
    		    10.times do |k|
    		      
		      data_row.set_format(k+5,data_format3)

		    end
		  end
                  data_row[i] = cus_column_content(columns[i],issue)
                end
              else 
                  data_row[i] = cus_column_content(columns[i], issue)
              end
            end

	    if !journals_hash[issue.id].nil?
	      journal = journals_hash[issue.id]
	        journal.visible_details.each do |detail|
	          if detail.property=="cf"
		    case detail.custom_field.name
		    when l(:shijiandian)
    			data_row.set_format(13,data_format4)

    		    when l(:shishituandui)
    			data_row.set_format(11,data_format4)
    		    
    		    when l(:fuzeren)
    			data_row.set_format(7,data_format4)
    		    
       		    when l(:pingtai)
    			data_row.set_format(8,data_format4)
    			
    		    when l(:yewubumen)
    			data_row.set_format(10,data_format4)
    		   
    		    when l(:beizhu)
    			data_row.set_format(14,data_format4)

    		    when l(:renwubianhao)
    			data_row.set_format(5,data_format4)
    			
    			
		    end
		  elsif detail.property=="attr"
		    case detail.prop_key
		    when 'create_issue'
		      9.times do |k|
			data_row.set_format(5+k,data_format5)
		      end

		    when 'status_id'
		        data_row.set_format(12,data_format4)
		    end
		    
		  end
	        end
	      end
	    end

        end


        book.write 'files/renwudabiao.xls'
        send_file('files/renwudabiao.xls')
      else
        respond_to do |format|
          format.html { render :template => 'issues/index', :layout => !request.xhr? }
          format.api  {
            Issue.load_visible_relations(@issues) if include_in_api_response?('relations')
          }
          format.atom { render_feed(@issues, :title => "#{@project || Setting.app_title}: #{l(:label_issue_plural)}") }
          format.csv  { send_data(query_to_csv(@issues, @query, params), :type => 'text/csv; header=present', :filename => 'issues.csv') }
          format.pdf  { send_data(issues_to_pdf(@issues, @project, @query), :type => 'application/pdf', :filename => 'issues.pdf') }
          
        end
      end
    else
      respond_to do |format|
        format.html { render(:template => 'issues/index', :layout => !request.xhr?) }
        format.any(:atom, :csv, :pdf) { render(:nothing => true) }
        format.api { render_validation_errors(@query) }
      end
    end
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  def show
    @journals = @issue.journals.includes(:user, :details, :issue).reorder("#{Journal.table_name}.id ASC").all
    @journals.each_with_index {|j,i| j.indice = i+1}
    @journals.reject!(&:private_notes?) unless User.current.allowed_to?(:view_private_notes, @issue.project)
    Journal.preload_journals_details_custom_fields(@journals)
    # TODO: use #select! when ruby1.8 support is dropped
    @journals.reject! {|journal| !journal.notes? && journal.visible_details.empty?}
    @journals.reverse! if User.current.wants_comments_in_reverse_order?

    @changesets = @issue.changesets.visible.all
    @changesets.reverse! if User.current.wants_comments_in_reverse_order?

    @relations = @issue.relations.select {|r| r.other_issue(@issue) && r.other_issue(@issue).visible? }
    @allowed_statuses = @issue.new_statuses_allowed_to(User.current)
    @edit_allowed = User.current.allowed_to?(:edit_issues, @project)
    @priorities = IssuePriority.active
    @time_entry = TimeEntry.new(:issue => @issue, :project => @issue.project)
    @relation = IssueRelation.new

    respond_to do |format|
      format.html {
        retrieve_previous_and_next_issue_ids
        render :template => 'issues/show'
      }
      format.api
      format.atom { render :template => 'journals/index', :layout => false, :content_type => 'application/atom+xml' }
      format.pdf  {
        pdf = issue_to_pdf(@issue, :journals => @journals)
        send_data(pdf, :type => 'application/pdf', :filename => "#{@project.identifier}-#{@issue.id}.pdf")
      }
    end
  end

  # Add a new issue
  # The new issue will be created from an existing one if copy_from parameter is given
  def new
    respond_to do |format|
      format.html { render :action => 'new', :layout => !request.xhr? }
    end
  end

  def create
    custom_field = CustomField.find_by_type_and_name("IssueCustomField",l(:renwubianhao))
    

    call_hook(:controller_issues_new_before_save, { :params => params, :issue => @issue })

   
    renwubianhao  = params[:issue]["custom_field_values"][custom_field.id.to_s]
    custom_value = CustomValue.find_by_customized_type_and_custom_field_id_and_value("Issue",custom_field.id,renwubianhao)
    if !custom_value.nil? && !custom_value.blank?
      flash[:notice] = l(:notice_renwubianhao_notunique)
      respond_to do |format|
        format.html { render :action => 'new' }
        format.api  { render_validation_errors(@issue) }
      end
    else
      @issue.save_attachments(params[:attachments] || (params[:issue] && params[:issue][:uploads]))  

      if @issue.save
        call_hook(:controller_issues_new_after_save, { :params => params, :issue => @issue})
        

        new_journal = Journal.new(              
                      :journalized_type => "Issue" ,
                      :journalized => @issue,
                      :user => User.current,
                      :notes => nil, 
                      :private_notes => false
        )
        new_journal.details = []
        new_journal.details << JournalDetail.new(:property=>'attr',
                                :prop_key=> 'create_issue',:old_value=>"",:value=>"")
        new_journal.notify = false
        new_journal.save


        respond_to do |format|
          format.html {
            render_attachment_warning_if_needed(@issue)
            flash[:notice] = l(:notice_issue_successful_create, :id => view_context.link_to("##{@issue.id}", issue_path(@issue), :title => @issue.subject))
            if params[:continue]
              attrs = {:tracker_id => @issue.tracker, :parent_issue_id => @issue.parent_issue_id}.reject {|k,v| v.nil?}
              redirect_to new_project_issue_path(@issue.project, :issue => attrs)
            else
              redirect_to issue_path(@issue)
            end
          }
          format.api  { render :action => 'show', :status => :created, :location => issue_url(@issue) }
        end
        return
      else
        respond_to do |format|
          format.html { render :action => 'new' }
          format.api  { render_validation_errors(@issue) }
        end
      end

    end

  end

  def edit
    return unless update_issue_from_params

    respond_to do |format|
      format.html { }
      format.xml  { }
    end
  end

  def update
    return unless update_issue_from_params
    @issue.save_attachments(params[:attachments] || (params[:issue] && params[:issue][:uploads]))
    saved = false
    begin
      saved = save_issue_with_child_records
    rescue ActiveRecord::StaleObjectError
      @conflict = true
      if params[:last_journal_id]
        @conflict_journals = @issue.journals_after(params[:last_journal_id]).all
        @conflict_journals.reject!(&:private_notes?) unless User.current.allowed_to?(:view_private_notes, @issue.project)
      end
    end

    if saved
      render_attachment_warning_if_needed(@issue)
      flash[:notice] = l(:notice_successful_update) unless @issue.current_journal.new_record?

      respond_to do |format|
        format.html { redirect_back_or_default issue_path(@issue) }
        format.api  { render_api_ok }
      end
    else
      respond_to do |format|
        format.html { render :action => 'edit' }
        format.api  { render_validation_errors(@issue) }
      end
    end
  end

  # Updates the issue form when changing the project, status or tracker
  # on issue creation/update
  def update_form
  end

  # Bulk edit/copy a set of issues
  def bulk_edit
    @issues.sort!
    @copy = params[:copy].present?
    @notes = params[:notes]

    if User.current.allowed_to?(:move_issues, @projects)
      @allowed_projects = Issue.allowed_target_projects_on_move
      if params[:issue]
        @target_project = @allowed_projects.detect {|p| p.id.to_s == params[:issue][:project_id].to_s}
        if @target_project
          target_projects = [@target_project]
        end
      end
    end
    target_projects ||= @projects

    if @copy
      @available_statuses = [IssueStatus.default]
    else
      @available_statuses = @issues.map(&:new_statuses_allowed_to).reduce(:&)
    end
    @custom_fields = target_projects.map{|p|p.all_issue_custom_fields.visible}.reduce(:&)
    @assignables = target_projects.map(&:assignable_users).reduce(:&)
    @trackers = target_projects.map(&:trackers).reduce(:&)
    @versions = target_projects.map {|p| p.shared_versions.open}.reduce(:&)
    @categories = target_projects.map {|p| p.issue_categories}.reduce(:&)
    if @copy
      @attachments_present = @issues.detect {|i| i.attachments.any?}.present?
      @subtasks_present = @issues.detect {|i| !i.leaf?}.present?
    end

    @safe_attributes = @issues.map(&:safe_attribute_names).reduce(:&)

    @issue_params = params[:issue] || {}
    @issue_params[:custom_field_values] ||= {}
  end

  def bulk_update
    @issues.sort!
    @copy = params[:copy].present?
    attributes = parse_params_for_bulk_issue_attributes(params)

    unsaved_issues = []
    saved_issues = []

    if @copy && params[:copy_subtasks].present?
      # Descendant issues will be copied with the parent task
      # Don't copy them twice
      @issues.reject! {|issue| @issues.detect {|other| issue.is_descendant_of?(other)}}
    end

    @issues.each do |orig_issue|
      orig_issue.reload
      if @copy
        issue = orig_issue.copy({},
          :attachments => params[:copy_attachments].present?,
          :subtasks => params[:copy_subtasks].present?
        )
      else
        issue = orig_issue
      end
      journal = issue.init_journal(User.current, params[:notes])
      issue.safe_attributes = attributes
      call_hook(:controller_issues_bulk_edit_before_save, { :params => params, :issue => issue })
      if issue.save
        saved_issues << issue
      else
        unsaved_issues << orig_issue
      end
    end

    if unsaved_issues.empty?
      flash[:notice] = l(:notice_successful_update) unless saved_issues.empty?
      if params[:follow]
        if @issues.size == 1 && saved_issues.size == 1
          redirect_to issue_path(saved_issues.first)
        elsif saved_issues.map(&:project).uniq.size == 1
          redirect_to project_issues_path(saved_issues.map(&:project).first)
        end
      else
        redirect_back_or_default _project_issues_path(@project)
      end
    else
      @saved_issues = @issues
      @unsaved_issues = unsaved_issues
      @issues = Issue.visible.where(:id => @unsaved_issues.map(&:id)).all
      bulk_edit
      render :action => 'bulk_edit'
    end
  end

  def destroy
    custom_field = CustomField.find_by_type_and_name("IssueCustomField",l(:renwubianhao))
    @hours = TimeEntry.where(:issue_id => @issues.map(&:id)).sum(:hours).to_f
    if @hours > 0
      case params[:todo]
      when 'destroy'
        # nothing to do
      when 'nullify'
        TimeEntry.where(['issue_id IN (?)', @issues]).update_all('issue_id = NULL')
      when 'reassign'
        reassign_to = @project.issues.find_by_id(params[:reassign_to_id])
        if reassign_to.nil?
          flash.now[:error] = l(:error_issue_not_found_in_project)
          return
        else
          TimeEntry.where(['issue_id IN (?)', @issues]).
            update_all("issue_id = #{reassign_to.id}")
        end
      else
        # display the destroy form if it's a user request
        return unless api_request?
      end
    end
    @issues.each do |issue|
      custom_value=""
      if issue.tracker.name==l(:renwu)
        custom_value = CustomValue.find_by_customized_id_and_custom_field_id(issue.id,custom_field.id)
      end

      begin
        issue.reload.destroy
      rescue ::ActiveRecord::RecordNotFound # raised by #reload if issue no longer exists
        # nothing to do, issue was already deleted (eg. by a parent)
      end

      if issue.tracker.name==l(:renwu)
        logger.info("***********************")
        logger.info("**$$"+issue.id.to_s+"**$$"+custom_field.id.to_s+"**$$"+custom_value.value)
        logger.info("*************************")
        
        new_journal = Journal.new(              
                    :journalized_type => "Issue" ,
                    :journalized => issue,
                    :user => User.current,
                    :notes => nil, 
                    :private_notes => false
        )
        new_journal.details = []
        new_journal.details << JournalDetail.new(:property=>'attr',
                                :prop_key=> 'destroy_issue',:old_value=> custom_value.value,:value=>issue.subject)
        new_journal.notify = false
        new_journal.save
      end
    end

    

    respond_to do |format|
      format.html { redirect_back_or_default _project_issues_path(@project) }
      format.api  { render_api_ok }
    end
  end

  private

  def find_project
    project_id = params[:project_id] || (params[:issue] && params[:issue][:project_id])
    @project = Project.find(project_id)
  rescue ActiveRecord::RecordNotFound
    render_404
  end

  def retrieve_previous_and_next_issue_ids
    retrieve_query_from_session
    if @query
      sort_init(@query.sort_criteria.empty? ? [['id', 'desc']] : @query.sort_criteria)
      sort_update(@query.sortable_columns, 'issues_index_sort')
      limit = 500
      issue_ids = @query.issue_ids(:order => sort_clause, :limit => (limit + 1), :include => [:assigned_to, :tracker, :priority, :category, :fixed_version])
      if (idx = issue_ids.index(@issue.id)) && idx < limit
        if issue_ids.size < 500
          @issue_position = idx + 1
          @issue_count = issue_ids.size
        end
        @prev_issue_id = issue_ids[idx - 1] if idx > 0
        @next_issue_id = issue_ids[idx + 1] if idx < (issue_ids.size - 1)
      end
    end
  end

  # Used by #edit and #update to set some common instance variables
  # from the params
  # TODO: Refactor, not everything in here is needed by #edit
  def update_issue_from_params
    @edit_allowed = User.current.allowed_to?(:edit_issues, @project)
    @time_entry = TimeEntry.new(:issue => @issue, :project => @issue.project)
    @time_entry.attributes = params[:time_entry]

    @issue.init_journal(User.current)

    issue_attributes = params[:issue]
    if issue_attributes && params[:conflict_resolution]
      case params[:conflict_resolution]
      when 'overwrite'
        issue_attributes = issue_attributes.dup
        issue_attributes.delete(:lock_version)
      when 'add_notes'
        issue_attributes = issue_attributes.slice(:notes)
      when 'cancel'
        redirect_to issue_path(@issue)
        return false
      end
    end
    @issue.safe_attributes = issue_attributes
    @priorities = IssuePriority.active
    @allowed_statuses = @issue.new_statuses_allowed_to(User.current)
    true
  end

  # TODO: Refactor, lots of extra code in here
  # TODO: Changing tracker on an existing issue should not trigger this
  def build_new_issue_from_params
    if params[:id].blank?
      @issue = Issue.new
      if params[:copy_from]
        begin
          @copy_from = Issue.visible.find(params[:copy_from])
          @copy_attachments = params[:copy_attachments].present? || request.get?
          @copy_subtasks = params[:copy_subtasks].present? || request.get?
          @issue.copy_from(@copy_from, :attachments => @copy_attachments, :subtasks => @copy_subtasks)
        rescue ActiveRecord::RecordNotFound
          render_404
          return
        end
      end
      @issue.project = @project
    else
      @issue = @project.issues.visible.find(params[:id])
    end

    @issue.project = @project
    @issue.author ||= User.current
    # Tracker must be set before custom field values
    @issue.tracker ||= @project.trackers.find((params[:issue] && params[:issue][:tracker_id]) || params[:tracker_id] || :first)
    if @issue.tracker.nil?
      render_error l(:error_no_tracker_in_project)
      return false
    end
    @issue.start_date ||= Date.today if Setting.default_issue_start_date_to_creation_date?
    @issue.safe_attributes = params[:issue]

    @priorities = IssuePriority.active
    @allowed_statuses = @issue.new_statuses_allowed_to(User.current, @issue.new_record?)
    @available_watchers = @issue.watcher_users
    if @issue.project.users.count <= 20
      @available_watchers = (@available_watchers + @issue.project.users.sort).uniq
    end
  end

  def check_for_default_issue_status
    if IssueStatus.default.nil?
      render_error l(:error_no_default_issue_status)
      return false
    end
  end

  def parse_params_for_bulk_issue_attributes(params)
    attributes = (params[:issue] || {}).reject {|k,v| v.blank?}
    attributes.keys.each {|k| attributes[k] = '' if attributes[k] == 'none'}
    if custom = attributes[:custom_field_values]
      custom.reject! {|k,v| v.blank?}
      custom.keys.each do |k|
        if custom[k].is_a?(Array)
          custom[k] << '' if custom[k].delete('__none__')
        else
          custom[k] = '' if custom[k] == '__none__'
        end
      end
    end
    attributes
  end

  # Saves @issue and a time_entry from the parameters
  def save_issue_with_child_records
    Issue.transaction do
      if params[:time_entry] && (params[:time_entry][:hours].present? || params[:time_entry][:comments].present?) && User.current.allowed_to?(:log_time, @issue.project)
        time_entry = @time_entry || TimeEntry.new
        time_entry.project = @issue.project
        time_entry.issue = @issue
        time_entry.user = User.current
        time_entry.spent_on = User.current.today
        time_entry.attributes = params[:time_entry]
        @issue.time_entries << time_entry
      end

      call_hook(:controller_issues_edit_before_save, { :params => params, :issue => @issue, :time_entry => time_entry, :journal => @issue.current_journal})
      if @issue.save
        call_hook(:controller_issues_edit_after_save, { :params => params, :issue => @issue, :time_entry => time_entry, :journal => @issue.current_journal})
      else
        raise ActiveRecord::Rollback
      end
    end
  end
end

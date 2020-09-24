class MeetingsController < ApplicationController
  include GraphHelper

  def create
    @new_event = create_event(access_token, create_params)

    redirect_to calendar_path
  end

  private

  def create_params
    params.require(:meeting).permit(:meeting_name, :start_time, :finish_time, :attendee_1_name, :attendee_1_email, :attendee_2_name, :attendee_2_email, :attendee_3_name, :attendee_3_email)
  end
end
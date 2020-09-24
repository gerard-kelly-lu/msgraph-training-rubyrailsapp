# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

require 'httparty'

# Graph API helper methods
module GraphHelper
  GRAPH_HOST = 'https://graph.microsoft.com'.freeze

  def make_api_call(endpoint, token, params = nil)
    headers = {
      Authorization: "Bearer #{token}"
    }

    query = params || {}

    HTTParty.get "#{GRAPH_HOST}#{endpoint}",
                 headers: headers,
                 query: query
  end

  # <GetCalendarSnippet>
  def get_calendar_events(token)
    get_events_url = '/v1.0/me/events'

    query = {
      '$select': 'subject,organizer,start,end,onlineMeeting',
      '$orderby': 'createdDateTime DESC'
    }

    response = make_api_call get_events_url, token, query

    raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 200

    response.parsed_response['value']
  end
  # </GetCalendarSnippet>

  def create_event(token, params)
    create_event_url = '/v1.0/me/events'

    headers = {
        Authorization: "Bearer #{token}",
        'Content-Type': 'application/json'
    }

    body = {
        "subject": params[:meeting_name],
        "body": {
            "contentType": "HTML",
            "content": "Does this time work for you?"
        },
        "start": {
            "dateTime": params[:start_time],
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": params[:finish_time],
            "timeZone": "UTC"
        },
        "location": {
            "displayName": "Cordova conference room"
        },
        "allowNewTimeProposals": true,
        "isOnlineMeeting": true,
        "onlineMeetingProvider": "teamsForBusiness",
        "attendees": build_attendees(params),
        "organizer": {
            "emailAddress": {
                "name": @user_name,
                "address": @user_email
            }
        }
    }.to_json

    response = HTTParty.post "#{GRAPH_HOST}#{create_event_url}",
                             headers: headers,
                             body: body

    raise response.parsed_response.to_s || "Request returned #{response.code}" unless response.code == 201

    response.parsed_response
  end

  private

  def build_attendees(params)
    attendees = []

    if params[:attendee_1_email]
      attendees << {
          "type": "required",
          "status": {
              "response": "none",
              "time": "0001-01-01T00:00:00Z"
          },
          "emailAddress": {
              "name": params[:attendee_1_name],
              "address": params[:attendee_1_email]
          }
      }
    end

    if params[:attendee_2_email]
      attendees << {
          "type": "required",
          "status": {
              "response": "none",
              "time": "0001-01-01T00:00:00Z"
          },
          "emailAddress": {
              "name": params[:attendee_2_name],
              "address": params[:attendee_2_email]
          }
      }
    end

    if params[:attendee_3_email]
      attendees << {
          "type": "required",
          "status": {
              "response": "none",
              "time": "0001-01-01T00:00:00Z"
          },
          "emailAddress": {
              "name": params[:attendee_3_name],
              "address": params[:attendee_3_email]
          }
      }
    end

    attendees
  end
end

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useRef, useState } from "react";
import { findIana } from "windows-iana";
import { Event } from "@microsoft/microsoft-graph-types";
import { AuthenticatedTemplate } from "@azure/msal-react";

import {
  CreateEvent,
  deleteEvent,
  getUserWeekCalendar,
  updateEvent,
} from "./GraphService";
import { useAppContext } from "./AppContext";
import "./Scheduler.css";
import "./App.css";

import {
  ScheduleComponent,
  Day,
  Week,
  WorkWeek,
  Month,
  Agenda,
  Inject,
} from "@syncfusion/ej2-react-schedule";

export default function Scheduler() {
  const app = useAppContext();

  const [events, setEvents] = useState<Event[]>();
  let scheduleObj = useRef<ScheduleComponent>(null);

  useEffect(() => {
    const loadEvents = async () => {
      if (app.user && !events) {
        try {
          debugger;
          let startDate: any = scheduleObj.current?.getCurrentViewDates()[0];
          let endDate: any = scheduleObj.current
            ?.getCurrentViewDates()
            .slice(-1)[0];
          const ianaTimeZones = findIana(app.user?.timeZone!);
          const events = await getUserWeekCalendar(
            app.authProvider!,
            ianaTimeZones[0].valueOf(),
            startDate,
            endDate
          );
          setEvents(events);
        } catch (err) {
          const error = err as Error;
          app.displayError!(error.message);
        }
      }
    };

    loadEvents();
  });

  const fieldsData = {
    id: "Id",
    subject: { name: "subject" },
    startTime: { name: "start" },
    endTime: { name: "end" },
  };

  const onActionComplete = async (args: any): Promise<void> => {
    let startDate: any;
    let endDate: any;
    if (args.requestType === "eventCreated") {
      const event = args.data[0];
      await CreateEvent(event);
    } else if (args.requestType === "eventChanged") {
      const event = args.data[0];
      await updateEvent(event);
    } else if (args.requestType === "eventRemoved") {
      var eventId = args.data[0].id;
      deleteEvent(eventId);
    }
    // Re-fetch events after any modification
    if (app.user) {
      const ianaTimeZones = findIana(app.user?.timeZone!);
      startDate = scheduleObj.current?.getCurrentViewDates()[0];
      endDate = scheduleObj.current?.getCurrentViewDates().slice(-1)[0];

      const updatedEvents = await getUserWeekCalendar(
        app.authProvider!,
        ianaTimeZones[0].valueOf(),
        startDate,
        endDate
      );
      setEvents(updatedEvents);
    }
  };

  const eventSettings = { dataSource: events, fields: fieldsData };

  return (
    <AuthenticatedTemplate>
      <ScheduleComponent
        ref={scheduleObj}
        eventSettings={eventSettings}
        height="100vh"
        actionComplete={onActionComplete}
      >
        <Inject services={[Day, Week, WorkWeek, Month, Agenda]} />
      </ScheduleComponent>
    </AuthenticatedTemplate>
  );
}

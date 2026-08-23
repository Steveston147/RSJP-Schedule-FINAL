'use client';

import dynamic from "next/dynamic";
import AuthControls from "../components/AuthControls";
import SessionBoundary from "../components/SessionBoundary";

const ScheduleApp = dynamic(() => import("../components/ScheduleApp"), {
  ssr: false,
});

export default function Page() {
  return (
    <SessionBoundary>
      <AuthControls />
      <ScheduleApp />
    </SessionBoundary>
  );
}

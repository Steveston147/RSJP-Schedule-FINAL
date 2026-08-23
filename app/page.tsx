'use client';

import dynamic from "next/dynamic";
import AuthControls from "../components/AuthControls";

const ScheduleApp = dynamic(() => import("../components/ScheduleApp"), {
  ssr: false,
});

export default function Page() {
  return (
    <>
      <AuthControls />
      <ScheduleApp />
    </>
  );
}

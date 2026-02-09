import dynamic from "next/dynamic";

const ScheduleApp = dynamic(() => import("../components/ScheduleApp"), {
  ssr: false,
});

export default function Page() {
  return <ScheduleApp />;
}

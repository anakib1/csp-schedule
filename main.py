from ortools.sat.python import cp_model
import pandas as pd


def create_schedule_one_week(groups, subjects, lecturers, rooms, days_per_week, slots_per_day):
    model = cp_model.CpModel()

    time_slots_per_week = days_per_week * slots_per_day
    time_slots = range(time_slots_per_week)

    X = {}
    for subject in subjects:
        for group in groups:
            if subject['name'] in group['subjects']:
                for lecturer in lecturers:
                    if subject['name'] in lecturer['subjects']:
                        for slot in time_slots:
                            X[(subject['name'], group['name'], slot, lecturer['name'])] = model.NewBoolVar(
                                f'X_{subject["name"]}_{group["name"]}_{slot}_{lecturer["name"]}'
                            )

    for room in rooms:
        for slot in time_slots:
            model.Add(
                sum(
                    X[(subject['name'], group['name'], slot, lecturer['name'])]
                    for subject in subjects
                    for group in groups
                    for lecturer in lecturers
                    if
                    (subject['name'], group['name'], slot, lecturer['name']) in X and room['size'] >= group['students']
                ) <= 1
            )

    for lecturer in lecturers:
        for slot in time_slots:
            model.Add(
                sum(
                    X[(subject['name'], group['name'], slot, lecturer['name'])]
                    for subject in subjects
                    for group in groups
                    if (subject['name'], group['name'], slot, lecturer['name']) in X
                ) <= 1
            )

    for group in groups:
        for slot in time_slots:
            model.Add(
                sum(
                    X[(subject['name'], group['name'], slot, lecturer['name'])]
                    for subject in subjects
                    for lecturer in lecturers
                    if (subject['name'], group['name'], slot, lecturer['name']) in X
                ) <= 1
            )

    deviation_vars = []
    for subject in subjects:
        for group in groups:
            if subject['name'] in group['subjects']:
                total_scheduled_hours = sum(
                    X[(subject['name'], group['name'], slot, lecturer['name'])]
                    for lecturer in lecturers
                    for slot in time_slots
                    if (subject['name'], group['name'], slot, lecturer['name']) in X
                )
                expected_hours_per_week = subject['hours'] // 21
                diff = model.NewIntVar(-time_slots_per_week, time_slots_per_week,
                                       f'diff_{subject["name"]}_{group["name"]}')
                deviation = model.NewIntVar(0, time_slots_per_week, f'deviation_{subject["name"]}_{group["name"]}')
                model.Add(diff == total_scheduled_hours - expected_hours_per_week)
                model.AddAbsEquality(deviation, diff)
                deviation_vars.append(deviation)

    model.Minimize(sum(deviation_vars))

    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        print("Optimal weekly schedule found:")
        weekly_schedule = []
        actual_hours = {group['name']: {subject['name']: 0 for subject in subjects} for group in groups}
        expected_hours = {group['name']: {subject['name']: subject['hours'] // 14 for subject in subjects if
                                          subject['name'] in group['subjects']} for group in groups}

        for slot in time_slots:
            for subject in subjects:
                for group in groups:
                    for lecturer in lecturers:
                        if (subject['name'], group['name'], slot, lecturer['name']) in X and \
                                solver.Value(X[(subject['name'], group['name'], slot, lecturer['name'])]) == 1:
                            day = slot // slots_per_day
                            time = slot % slots_per_day
                            weekly_schedule.append({
                                'day': day,
                                'time': time,
                                'subject': subject['name'],
                                'group': group['name'],
                                'lecturer': lecturer['name']
                            })
                            actual_hours[group['name']][subject['name']] += 1

        print("\nActual vs Expected Hours:")
        for group_name in actual_hours:
            print(f"Group {group_name}:")
            for subject_name, hours in actual_hours[group_name].items():
                expected = expected_hours[group_name].get(subject_name, 0)
                print(f"  {subject_name}: Actual = {hours}, Expected = {expected}")

        return weekly_schedule
    else:
        print("No feasible schedule found.")
        return None


groups_file = "groups.csv"
lecturers_file = "lecturers.csv"
rooms_file = "rooms.csv"
subjects_file = "subjects.csv"


def load_input_from_csv(groups_file, lecturers_file, rooms_file, subjects_file):
    groups_df = pd.read_csv(groups_file)
    groups = [
        {
            "name": row["name"],
            "students": row["num_students"],
            "subjects": row["subjects"].split(";")
        }
        for _, row in groups_df.iterrows()
    ]

    lecturers_df = pd.read_csv(lecturers_file)
    lecturers = [
        {
            "name": row["name"],
            "subjects": row["subjects_can_teach"].split(";")
        }
        for _, row in lecturers_df.iterrows()
    ]

    rooms_df = pd.read_csv(rooms_file)
    rooms = [
        {
            "name": row["name"],
            "size": row["capacity"]
        }
        for _, row in rooms_df.iterrows()
    ]

    subjects_df = pd.read_csv(subjects_file)
    subjects = [
        {
            "name": row["name"],
            "hours": row["total_hours"]
        }
        for _, row in subjects_df.iterrows()
    ]

    return groups, lecturers, rooms, subjects


def export_schedule_to_excel(schedule, days_per_week, slots_per_day, groups, filename="schedule.xlsx"):
    columns = [f"Day {day + 1} Slot {slot + 1}" for day in range(days_per_week) for slot in range(slots_per_day)]

    schedule_df = pd.DataFrame(index=[group['name'] for group in groups], columns=columns)

    for lecture in schedule:
        day = lecture['day']
        slot = lecture['time']
        column = f"Day {day + 1} Slot {slot + 1}"
        group = lecture['group']
        subject = lecture['subject']
        lecturer = lecture['lecturer']

        schedule_df.loc[group, column] = f"{subject} by {lecturer}"

    schedule_df.to_excel(filename, sheet_name="Weekly Schedule", index=True)
    print(f"Schedule exported to {filename}")


if __name__ == '__main__':
    groups, lecturers, rooms, subjects = load_input_from_csv(groups_file, lecturers_file, rooms_file, subjects_file)

    print("Groups:", groups)
    print("Lecturers:", lecturers)
    print("Rooms:", rooms)
    print("Subjects:", subjects)

    days_per_week = 5
    slots_per_day = 4

    schedule = create_schedule_one_week(groups, subjects, lecturers, rooms, days_per_week, slots_per_day)

    if schedule:
        print("Weekly Schedule:")
        for lecture in schedule:
            print(
                f"  Day {lecture['day'] + 1}, Time Slot {lecture['time'] + 1}: {lecture['subject']} for {lecture['group']} by {lecture['lecturer']}")

    export_schedule_to_excel(schedule, days_per_week, slots_per_day, groups)

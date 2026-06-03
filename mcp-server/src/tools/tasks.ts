import { RegisteredTool, ok, fail, accountFields } from "./toolTypes.js";
import { hasTasks } from "../types/graph.js";

export const taskTools: RegisteredTool[] = [
  {
    name: "list_task_lists",
    description: "List the user's Microsoft To Do task lists. Use a returned list ID with the task tools.",
    inputSchema: {
      type: "object",
      properties: { ...accountFields },
    },
    async run(provider) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      return ok(await provider.listTaskLists());
    },
  },
  {
    name: "list_tasks",
    description: "List tasks in a Microsoft To Do list. Excludes completed tasks unless includeCompleted is true.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID (from list_task_lists)" },
        includeCompleted: { type: "boolean", description: "Include completed tasks" },
        top: { type: "number", description: "Max results (default 50)" },
        ...accountFields,
      },
      required: ["listId"],
    },
    async run(provider, args) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      return ok(
        await provider.listTasks(args.listId as string, {
          includeCompleted: args.includeCompleted as boolean | undefined,
          top: Number(args.top) || undefined,
        })
      );
    },
  },
  {
    name: "create_task",
    description: "Create a task in a Microsoft To Do list.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        title: { type: "string", description: "Task title" },
        body: { type: "string", description: "Notes / details" },
        dueDateTime: { type: "string", description: "ISO-8601 due datetime (UTC)" },
        reminderDateTime: { type: "string", description: "ISO-8601 reminder datetime (UTC)" },
        importance: { type: "string", enum: ["low", "normal", "high"] },
        ...accountFields,
      },
      required: ["listId", "title"],
    },
    async run(provider, args) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      const task = await provider.createTask(args.listId as string, {
        title: args.title as string,
        body: args.body as string | undefined,
        dueDateTime: args.dueDateTime as string | undefined,
        reminderDateTime: args.reminderDateTime as string | undefined,
        importance: args.importance as "low" | "normal" | "high" | undefined,
      });
      return ok({ success: true, task });
    },
  },
  {
    name: "update_task",
    description: "Update a task's title, notes, due date, importance, or status.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
        title: { type: "string" },
        body: { type: "string" },
        dueDateTime: { type: "string", description: "ISO-8601 due datetime (UTC)" },
        importance: { type: "string", enum: ["low", "normal", "high"] },
        status: { type: "string", enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"] },
        ...accountFields,
      },
      required: ["listId", "taskId"],
    },
    async run(provider, args) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      const { listId, taskId, accountId, mailbox, ...changes } = args as Record<string, unknown>;
      const task = await provider.updateTask(listId as string, taskId as string, changes as any);
      return ok({ success: true, task });
    },
  },
  {
    name: "complete_task",
    description: "Mark a Microsoft To Do task as completed.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
        ...accountFields,
      },
      required: ["listId", "taskId"],
    },
    async run(provider, args) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      const task = await provider.completeTask(args.listId as string, args.taskId as string);
      return ok({ success: true, task });
    },
  },
  {
    name: "delete_task",
    description: "Delete a task from a Microsoft To Do list.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "Task list ID" },
        taskId: { type: "string", description: "Task ID" },
        ...accountFields,
      },
      required: ["listId", "taskId"],
    },
    async run(provider, args) {
      if (!hasTasks(provider)) return fail("Tasks are not supported for this account.");
      await provider.deleteTask(args.listId as string, args.taskId as string);
      return ok({ success: true });
    },
  },
];

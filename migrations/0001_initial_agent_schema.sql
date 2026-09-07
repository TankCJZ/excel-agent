-- Excel Agent Standalone D1 Schema Migration
-- Defines tables for conversations, messages, tasks, workbooks, rate limits, and mutations.

CREATE TABLE IF NOT EXISTS `users` (
	`id` text PRIMARY KEY NOT NULL,
	`email` text NOT NULL,
	`password_hash` text NOT NULL,
	`display_name` text,
	`credits_balance` integer DEFAULT 0 NOT NULL,
	`free_credits_balance` integer DEFAULT 0 NOT NULL,
	`subscription_credits_balance` integer DEFAULT 0 NOT NULL,
	`free_xlsx_download_task_id` text,
	`first_login_bonus_granted_at` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS `users_email_unique` ON `users` (`email`);

CREATE TABLE IF NOT EXISTS `subscriptions` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`stripe_customer_id` text NOT NULL,
	`stripe_subscription_id` text NOT NULL,
	`stripe_price_id` text,
	`plan_key` text NOT NULL,
	`interval` text NOT NULL,
	`version` integer NOT NULL,
	`status` text NOT NULL,
	`current_period_start` text,
	`current_period_end` text,
	`cancel_at_period_end` integer DEFAULT 0 NOT NULL,
	`last_credited_invoice_id` text,
	`last_credited_period_start` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	`canceled_at` text,
	FOREIGN KEY (`user_id`) REFERENCES `users`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE INDEX IF NOT EXISTS `subscriptions_user_id_idx` ON `subscriptions` (`user_id`);
CREATE UNIQUE INDEX IF NOT EXISTS `subscriptions_stripe_subscription_id_unique` ON `subscriptions` (`stripe_subscription_id`);

CREATE TABLE IF NOT EXISTS `credit_ledger` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`task_id` text,
	`type` text NOT NULL,
	`amount` integer NOT NULL,
	`balance_after` integer NOT NULL,
	`idempotency_key` text NOT NULL,
	`note` text,
	`created_at` text NOT NULL,
	FOREIGN KEY (`user_id`) REFERENCES `users`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE INDEX IF NOT EXISTS `credit_ledger_user_id_idx` ON `credit_ledger` (`user_id`);
CREATE INDEX IF NOT EXISTS `credit_ledger_task_id_idx` ON `credit_ledger` (`task_id`);
CREATE UNIQUE INDEX IF NOT EXISTS `credit_ledger_idempotency_key_unique` ON `credit_ledger` (`idempotency_key`);

CREATE TABLE IF NOT EXISTS `excel_agent_workbooks` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text DEFAULT 'legacy-demo' NOT NULL,
	`session_id` text NOT NULL,
	`r2_key` text NOT NULL,
	`file_name` text NOT NULL,
	`size_bytes` integer NOT NULL,
	`summary_json` text NOT NULL,
	`context_r2_key` text,
	`source_etag` text,
	`context_version` integer,
	`parsed_at` text,
	`parent_workbook_id` text,
	`root_workbook_id` text,
	`source_task_id` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL
);
CREATE INDEX IF NOT EXISTS `excel_agent_workbooks_user_id_idx` ON `excel_agent_workbooks` (`user_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_workbooks_owner_session_idx` ON `excel_agent_workbooks` (`user_id`,`session_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_workbooks_session_id_idx` ON `excel_agent_workbooks` (`session_id`);
CREATE UNIQUE INDEX IF NOT EXISTS `excel_agent_workbooks_r2_key_unique` ON `excel_agent_workbooks` (`r2_key`);
CREATE INDEX IF NOT EXISTS `excel_agent_workbooks_created_at_idx` ON `excel_agent_workbooks` (`created_at`);

CREATE TABLE IF NOT EXISTS `excel_agent_conversations` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`session_id` text NOT NULL,
	`title` text DEFAULT '' NOT NULL,
	`workbook_id` text,
	`workbook_revision` integer DEFAULT 0 NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	`last_message_at` text
);
CREATE INDEX IF NOT EXISTS `excel_agent_conversations_user_id_idx` ON `excel_agent_conversations` (`user_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_conversations_user_updated_idx` ON `excel_agent_conversations` (`user_id`,`updated_at`);
CREATE INDEX IF NOT EXISTS `excel_agent_conversations_session_id_idx` ON `excel_agent_conversations` (`session_id`);

CREATE TABLE IF NOT EXISTS `excel_agent_conversation_tombstones` (
	`conversation_id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`session_id` text NOT NULL,
	`deleted_at` text NOT NULL
);

CREATE TABLE IF NOT EXISTS `excel_agent_conversation_workbooks` (
	`conversation_id` text NOT NULL,
	`workbook_id` text NOT NULL,
	`added_at` text NOT NULL,
	PRIMARY KEY (`conversation_id`, `workbook_id`),
	FOREIGN KEY (`conversation_id`) REFERENCES `excel_agent_conversations`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`workbook_id`) REFERENCES `excel_agent_workbooks`(`id`) ON UPDATE no action ON DELETE cascade
);

CREATE TABLE IF NOT EXISTS `excel_agent_tasks` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text DEFAULT 'legacy-demo' NOT NULL,
	`session_id` text NOT NULL,
	`thread_id` text NOT NULL,
	`task_type` text DEFAULT 'workbook_export' NOT NULL,
	`workflow_instance_id` text,
	`input_r2_key` text NOT NULL,
	`input_name` text NOT NULL,
	`output_r2_key` text,
	`output_name` text,
	`prompt` text NOT NULL,
	`plan_json` text,
	`result_json` text,
	`status` text NOT NULL,
	`error_message` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	`completed_at` text
);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_user_id_idx` ON `excel_agent_tasks` (`user_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_owner_thread_idx` ON `excel_agent_tasks` (`user_id`,`thread_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_session_id_idx` ON `excel_agent_tasks` (`session_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_thread_id_idx` ON `excel_agent_tasks` (`thread_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_status_idx` ON `excel_agent_tasks` (`status`);
CREATE INDEX IF NOT EXISTS `excel_agent_tasks_created_at_idx` ON `excel_agent_tasks` (`created_at`);

CREATE TABLE IF NOT EXISTS `excel_agent_messages` (
	`id` text PRIMARY KEY NOT NULL,
	`conversation_id` text NOT NULL,
	`turn_id` text NOT NULL,
	`role` text NOT NULL,
	`status` text NOT NULL,
	`content` text DEFAULT '' NOT NULL,
	`task_id` text,
	`workbook_id` text,
	`metadata_json` text,
	`error_code` text,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`conversation_id`) REFERENCES `excel_agent_conversations`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE INDEX IF NOT EXISTS `excel_agent_messages_conversation_id_idx` ON `excel_agent_messages` (`conversation_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_messages_conversation_created_idx` ON `excel_agent_messages` (`conversation_id`,`created_at`);
CREATE INDEX IF NOT EXISTS `excel_agent_messages_task_id_idx` ON `excel_agent_messages` (`task_id`);
CREATE UNIQUE INDEX IF NOT EXISTS `excel_agent_messages_turn_role_unique` ON `excel_agent_messages` (`conversation_id`,`turn_id`,`role`);

CREATE TABLE IF NOT EXISTS `excel_agent_task_events` (
	`id` text PRIMARY KEY NOT NULL,
	`task_id` text NOT NULL,
	`type` text NOT NULL,
	`message` text NOT NULL,
	`metadata_json` text,
	`created_at` text NOT NULL,
	FOREIGN KEY (`task_id`) REFERENCES `excel_agent_tasks`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE INDEX IF NOT EXISTS `excel_agent_task_events_task_id_idx` ON `excel_agent_task_events` (`task_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_task_events_created_at_idx` ON `excel_agent_task_events` (`created_at`);

CREATE TABLE IF NOT EXISTS `excel_agent_rate_limits` (
	`key` text PRIMARY KEY NOT NULL,
	`count` integer NOT NULL,
	`expires_at` text NOT NULL,
	`updated_at` text NOT NULL
);
CREATE INDEX IF NOT EXISTS `excel_agent_rate_limits_expires_at_idx` ON `excel_agent_rate_limits` (`expires_at`);

CREATE TABLE IF NOT EXISTS `excel_agent_mutations` (
	`task_id` text PRIMARY KEY NOT NULL,
	`conversation_id` text NOT NULL,
	`turn_id` text NOT NULL,
	`payload_json` text NOT NULL,
	`result_json` text,
	`state` text DEFAULT 'queued' NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`task_id`) REFERENCES `excel_agent_tasks`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE UNIQUE INDEX IF NOT EXISTS `excel_agent_mutations_conversation_turn_unique` ON `excel_agent_mutations` (`conversation_id`,`turn_id`);
CREATE INDEX IF NOT EXISTS `excel_agent_mutations_state_idx` ON `excel_agent_mutations` (`state`,`updated_at`);

CREATE TABLE IF NOT EXISTS `credit_reservations` (
	`id` text PRIMARY KEY NOT NULL,
	`user_id` text NOT NULL,
	`conversation_id` text NOT NULL,
	`turn_id` text NOT NULL,
	`task_id` text,
	`policy_version` integer DEFAULT 1 NOT NULL,
	`estimated_credits` integer NOT NULL,
	`held_free_credits` integer DEFAULT 0 NOT NULL,
	`held_subscription_credits` integer DEFAULT 0 NOT NULL,
	`settled_credits` integer DEFAULT 0 NOT NULL,
	`status` text DEFAULT 'reserved' NOT NULL,
	`expires_at` text NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	`settled_at` text,
	`released_at` text,
	`finalization_key` text
);
CREATE INDEX IF NOT EXISTS `credit_reservations_user_id_idx` ON `credit_reservations` (`user_id`);
CREATE INDEX IF NOT EXISTS `credit_reservations_task_id_idx` ON `credit_reservations` (`task_id`);
CREATE INDEX IF NOT EXISTS `credit_reservations_status_expires_idx` ON `credit_reservations` (`status`,`expires_at`);
CREATE UNIQUE INDEX IF NOT EXISTS `credit_reservations_turn_id_unique` ON `credit_reservations` (`turn_id`);

CREATE TABLE IF NOT EXISTS `credit_usage_events` (
	`id` text PRIMARY KEY NOT NULL,
	`reservation_id` text NOT NULL,
	`user_id` text NOT NULL,
	`turn_id` text NOT NULL,
	`task_id` text,
	`action` text NOT NULL,
	`credits` integer NOT NULL,
	`metadata_json` text,
	`idempotency_key` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`reservation_id`) REFERENCES `credit_reservations`(`id`) ON UPDATE no action ON DELETE cascade
);
CREATE INDEX IF NOT EXISTS `credit_usage_events_reservation_id_idx` ON `credit_usage_events` (`reservation_id`);
CREATE INDEX IF NOT EXISTS `credit_usage_events_user_id_idx` ON `credit_usage_events` (`user_id`);
CREATE INDEX IF NOT EXISTS `credit_usage_events_turn_id_idx` ON `credit_usage_events` (`turn_id`);
CREATE UNIQUE INDEX IF NOT EXISTS `credit_usage_events_idempotency_key_unique` ON `credit_usage_events` (`idempotency_key`);

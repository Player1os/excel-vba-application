// Load local modules.
const {
	notify,
	success,
} = require('../../common/format')
const spawnChildProcess = require('../../common/spawn_child_process')

// Define the current task name.
const taskName = 'UPDATE BASE: CREATE'

;(async () => {

	// Fetch the base remote repository's branches.
	await spawnChildProcess.inherited(taskName, 'git', ['fetch', 'base'])

	// Make sure we are in the project's master branch.
	await spawnChildProcess.inherited(taskName, 'git', ['checkout', 'master'])

	// Rebase the project's master branch to the base remote repository's master branch.
	await spawnChildProcess.inherited(taskName, 'git', ['rebase', 'base/master'], true)

	// Loop until the rebase is finished or aborted.
	while ((await spawnChildProcess.piped(taskName, 'git', ['status'])).stdout.substr(0, 18) === 'rebase in progress') {
		// Notify the user of the instructions to follow.
		notify(taskName, 'follow the rebase instructions and when done enter the `exit` command, to continue.')

		// Execute the nested shell.
		await spawnChildProcess.inherited(taskName, 'cmd')
	}

	// Execute garbage collection upon the repository.
	await spawnChildProcess.inherited(taskName, 'git', ['gc'])

	// Propagate the changes to the origin remote repository.
	await spawnChildProcess.inherited(taskName, 'git', ['push', '-f'])
	await spawnChildProcess.inherited(taskName, 'git', ['push', '--tags', '-f'])

	// Report the task's success.
	success(taskName)
})()
